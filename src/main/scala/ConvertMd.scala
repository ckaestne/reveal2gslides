import com.google.api.client.json.gson.GsonFactory
import com.google.api.services.slides.v1.{Slides, SlidesScopes, model}
import com.google.auth.Credentials
import com.google.auth.http.HttpCredentialsAdapter
import com.google.auth.oauth2.GoogleCredentials
import com.google.auth.oauth2.ServiceAccountCredentials
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport
import com.google.api.client.http.HttpRequest
import com.google.api.client.json.gson.GsonFactory
import com.google.api.client.util.DateTime
import com.google.auth.Credentials
import com.google.auth.http.HttpCredentialsAdapter
import com.google.auth.oauth2.ServiceAccountCredentials

import scala.collection.immutable.List
import com.google.api.services.slides.v1.model.*

import java.io.{File, FileInputStream, FileReader, FileWriter}
import java.util.Collections
import scala.collection.JavaConverters.*

object ConvertMd {


  import org.commonmark.node.*;
  import org.commonmark.parser.Parser;
  import org.commonmark.renderer.html.HtmlRenderer;

  type Slide = List[Node]

  def readSlides(filepath: String): List[Slide] = {

    val parser = Parser.builder().build();
    val document = parser.parseReader(new FileReader(filepath))

    var node = document.getFirstChild
    var slides: List[List[Node]] = List()
    var currentSlide: List[Node] = List()
    while (node != null) {

      node match {
        case tb: ThematicBreak =>
          slides ::= currentSlide.reverse
          currentSlide = List()
        case _ =>
          currentSlide ::= node
      }

      node = node.getNext
    }

    slides ::= currentSlide.reverse
    slides.reverse


  }

  private def getHeadings(slide: Slide): List[Heading] = slide.collect({ case h: Heading => h })

  private def isMainHeadingSlide(slide: Slide): Boolean =
    getHeadings(slide).exists(_.getLevel == 1)

  case class TextContentContext(offset: Int, slideId: String, objectId: String /*of textfield*/) {
    def shift(offsetShift: Int): TextContentContext = TextContentContext(offset + offsetShift, slideId, objectId)
  }

  case class TextContent(text: String, formatting: TextContentContext => List[Request] = _ => Nil) {
    def :+(s: String): TextContent = TextContent(text + s, formatting)

    def +:(s: String): TextContent = TextContent(s + text, i => formatting(i shift s.length))

    def concat(other: TextContent): TextContent =
      TextContent(text + other.text, i => formatting(i) ++ other.formatting(i shift text.length))

    def setStyle(style: TextStyle, fields: String): TextContent =
      addFormatting(c => List(new Request().setUpdateTextStyle(new UpdateTextStyleRequest()
        .setObjectId(c.objectId)
        .setTextRange(new Range().setStartIndex(c.offset).setEndIndex(c.offset + text.length).setType("FIXED_RANGE"))
        .setStyle(style)
        .setFields(fields)
      )))

    def makeItalics(): TextContent = setStyle(new TextStyle().setItalic(true), "italic")

    def makeBold(): TextContent = setStyle(new TextStyle().setBold(true), "bold")

    def makeCode(): TextContent =
      setStyle(new TextStyle().setFontFamily("Consolas"), "fontFamily")
        .addFormatting(c => List(new Request().setUpdateParagraphStyle(new UpdateParagraphStyleRequest()
          .setObjectId(c.objectId)
          .setTextRange(new Range().setStartIndex(c.offset).setEndIndex(c.offset + text.length).setType("FIXED_RANGE"))
          .setStyle(new ParagraphStyle().setSpaceBelow(new Dimension().setMagnitude(0).setUnit("PT")))
          .setFields("spaceBelow")
        )))

    def makeTodo(): TextContent = setStyle(new TextStyle()
      .setBold(true)
      .setForegroundColor(new OptionalColor().setOpaqueColor(new OpaqueColor().setRgbColor(new RgbColor().setRed(1.0f))))
      , "bold,foregroundColor")

    def makeItemize(): TextContent = addFormatting(c => List(new Request().setCreateParagraphBullets(new CreateParagraphBulletsRequest()
      .setObjectId(c.objectId)
      .setTextRange(new Range().setStartIndex(c.offset).setEndIndex(c.offset + text.length).setType("FIXED_RANGE")))
    ))

    def makeEnumerate(): TextContent = addFormatting(c => List(new Request().setCreateParagraphBullets(new CreateParagraphBulletsRequest()
      .setObjectId(c.objectId)
      .setTextRange(new Range().setStartIndex(c.offset).setEndIndex(c.offset + text.length).setType("FIXED_RANGE"))
      .setBulletPreset("NUMBERED_DIGIT_ALPHA_ROMAN"))
    ))

    def makeIndent(): TextContent = addFormatting(c => List(new Request().setUpdateParagraphStyle(new UpdateParagraphStyleRequest()
      .setObjectId(c.objectId)
      .setTextRange(new Range().setStartIndex(c.offset).setEndIndex(c.offset + text.length).setType("FIXED_RANGE"))
      .setFields("indentStart,indentFirstLine")
      .setStyle(new ParagraphStyle().setIndentFirstLine(new Dimension().setMagnitude(36).setUnit("PT")).setIndentStart(new Dimension().setMagnitude(36).setUnit("PT")))
    )))

    def addFormatting(f: TextContentContext => List[Request]): TextContent = TextContent(text, c => formatting(c) ++ f(c))

    def genTextAndFormattingRequests(context: TextContentContext): List[Request] = if (text.isEmpty) Nil else
      new Request().setInsertText(new InsertTextRequest().setText(text).setObjectId(context.objectId)) +: formatting(context)

    //    override def toString: String = ???
  }

  def concat(texts: Seq[TextContent]): TextContent = if (texts.isEmpty) TextContent("") else texts.reduce(_ concat _)

  def matchHtmlComment(name: String): Node => Boolean = {
    case html: HtmlBlock => html.getLiteral.matches(s"<!--\\W*$name\\W*-->")
    case _ => false
  }

  def checkNoContentBeforeHeading(s: Slide): Unit = {
    val contentBeforeHeading = s.takeWhile(!_.isInstanceOf[Heading]).filterNot(_.isInstanceOf[HtmlBlock])
    val firstHeading = s.find(_.isInstanceOf[Heading])
    assert(firstHeading.isEmpty || contentBeforeHeading.isEmpty, s"content before heading not supported $firstHeading")
  }

  def filterLayoutInstructions(s: Slide): Slide = s.filterNot {
    case html: HtmlBlock => Set("col", "colstart", "colend").exists(s => matchHtmlComment(s)(html))
    case html: HtmlBlock if html.getLiteral.startsWith("<div class=\"small") => true
    case html: HtmlBlock if html.getLiteral.startsWith("<div class=\"stretch") => true
    case html: HtmlBlock if html.getLiteral == "</div>" => true
    case _ => false
  }

  def convertSlide(slide: Slide, slideId: String): (Seq[Request], Seq[Presentation => Seq[Request]]) = {
    val notesNodeIdx = slide.indexWhere(n => n.isInstanceOf[Text] && n.asInstanceOf[Text].getLiteral.startsWith("Note:"))
    val (main, notes) = slide.splitAt(if (notesNodeIdx < 0) Int.MaxValue else notesNodeIdx)

    val referencesNodeIdx = main.indexWhere(matchHtmlComment("references?_?"))
    val (main2, references) = main.splitAt(if (referencesNodeIdx < 0) Int.MaxValue else referencesNodeIdx)

    val colStartIdx = main2.indexWhere(matchHtmlComment("colstart"))
    assert(main2.count(matchHtmlComment("col")) <= 1, "at most 2 columns per slide supported")
    val colIdx = main2.indexWhere(matchHtmlComment("col"))
    val colEndIdx = main2.indexWhere(matchHtmlComment("colend"))
    val contentBeforeColStart = main2.take(colStartIdx).filterNot(_.isInstanceOf[HtmlBlock]).nonEmpty
    val onlyHeaderBeforeColStart = main2.take(colStartIdx).filterNot(_.isInstanceOf[HtmlBlock]).filterNot(_.isInstanceOf[Heading]).isEmpty
    val contentAfterColEnd = main2.drop(colEndIdx).filterNot(_.isInstanceOf[HtmlBlock]).nonEmpty
    assert(colStartIdx <= colIdx && colIdx <= colEndIdx, "colstart, col, colend must be in order")


    //heading and slide body (other than speaker notes and references)
    var result =
      if (colStartIdx < 0) {
        checkNoContentBeforeHeading(main2)
        convertStandardSlide(filterLayoutInstructions(main2), filterLayoutInstructions(references), slideId)
      } else if ((!contentBeforeColStart || onlyHeaderBeforeColStart) && !contentAfterColEnd) {
        val (rest, right) = main2.splitAt(colIdx)
        val (heading, left) = rest.splitAt(colStartIdx)
        checkNoContentBeforeHeading(heading)
        convertTwoColumnLayout(filterLayoutInstructions(heading), filterLayoutInstructions(left), filterLayoutInstructions(right), filterLayoutInstructions(references), slideId)
      } else {
        val (aboveBottom, bottom) = main2.splitAt(colEndIdx)
        val (beforeRight, right) = aboveBottom.splitAt(colIdx)
        val (top, left) = beforeRight.splitAt(colStartIdx)
        checkNoContentBeforeHeading(top)
        convertFourPartLayout(filterLayoutInstructions(top), filterLayoutInstructions(left), filterLayoutInstructions(right), filterLayoutInstructions(bottom), filterLayoutInstructions(references), slideId)
      }


    //speaker notes
    result = (result._1, result._2 ++ convertSpeakerNotes(notes, slideId))


    result
  }

  def convertSpeakerNotes(notes: List[Node], slideId: String): Seq[Presentation => Seq[Request]] =
    //TODO
    createNote(slideId, TextContent("Note")) :: Nil

  def convertStandardSlide(slide: Slide, references: Slide, slideId: String): (Seq[Request], Seq[Presentation => Seq[Request]]) = {
    var requests: Seq[Request] = Nil
    requests :+= new Request().setCreateSlide(new CreateSlideRequest().setObjectId(slideId)
      .setSlideLayoutReference(
        new LayoutReference().setPredefinedLayout("TITLE_AND_BODY"))
      .setPlaceholderIdMappings(List(
        new LayoutPlaceholderIdMapping().setObjectId(slideId + "_title").setLayoutPlaceholder(
          new Placeholder().setType("TITLE").setIndex(0)),
        new LayoutPlaceholderIdMapping().setObjectId(slideId + "_body").setLayoutPlaceholder(
          new Placeholder().setType("BODY").setIndex(0))
      ).asJava))

    val headings = getHeadings(slide)
    assert(headings.size <= 1, s"only one heading per slide supported ${headings.map(getSlideTitle)}")
    headings.foreach(heading => {
      requests :+= new Request().setInsertText(new InsertTextRequest().setObjectId(slideId + "_title").setText(getSlideTitle(heading)))
    })

    requests ++= convertContent(filterHeadings(slide)).genTextAndFormattingRequests(TextContentContext(0, slideId, slideId + "_body"))

    requests ++= createFootnote(Nil, slideId, List(slideId + "_body"))

    (requests, Nil)
  }

  def convertTwoColumnLayout(heading: List[Node], left: List[Node], right: List[Node], references: List[Node], slideId: String): (Seq[Request], Seq[Presentation => Seq[Request]]) = {
    var requests: Seq[Request] = Nil
    requests :+= new Request().setCreateSlide(new CreateSlideRequest().setObjectId(slideId)
      .setSlideLayoutReference(
        new LayoutReference().setPredefinedLayout("TITLE_AND_TWO_COLUMNS"))
      .setPlaceholderIdMappings(List(
        new LayoutPlaceholderIdMapping().setObjectId(slideId + "_title").setLayoutPlaceholder(
          new Placeholder().setType("TITLE").setIndex(0)),
        new LayoutPlaceholderIdMapping().setObjectId(slideId + "_left").setLayoutPlaceholder(
          new Placeholder().setType("BODY").setIndex(0)),
        new LayoutPlaceholderIdMapping().setObjectId(slideId + "_right").setLayoutPlaceholder(
          new Placeholder().setType("BODY").setIndex(1))
      ).asJava))

    val headings = getHeadings(heading)
    assert(headings.size <= 1, "only one heading per slide supported")
    headings.foreach(heading => {
      requests :+= new Request().setInsertText(new InsertTextRequest().setObjectId(slideId + "_title").setText(getSlideTitle(heading)))
    })

    requests ++= convertContent(left, true).genTextAndFormattingRequests(TextContentContext(0, slideId, slideId + "_left"))
    requests ++= convertContent(right, true).genTextAndFormattingRequests(TextContentContext(0, slideId, slideId + "_right"))

    requests ++= createFootnote(references, slideId, List(slideId + "_left", slideId + "_right"))

    (requests, Nil)
  }


  def convertFourPartLayout(top: List[Node], left: List[Node], right: List[Node], bottom: List[Node], references: List[Node], slideId: String): (Seq[Request], Seq[Presentation => Seq[Request]]) = {
    var requests: Seq[Request] = Nil
    val titleId = slideId + "_title"
    val topId = slideId + "_top"
    val bottomId = slideId + "_bottom"
    val leftId = slideId + "_left"
    val rightId = slideId + "_right"

    requests :+= new Request().setCreateSlide(new CreateSlideRequest().setObjectId(slideId)
      .setSlideLayoutReference(
        new LayoutReference().setPredefinedLayout("TITLE_AND_TWO_COLUMNS"))
      .setPlaceholderIdMappings(List(
        new LayoutPlaceholderIdMapping().setObjectId(titleId).setLayoutPlaceholder(
          new Placeholder().setType("TITLE").setIndex(0)),
        new LayoutPlaceholderIdMapping().setObjectId(leftId).setLayoutPlaceholder(
          new Placeholder().setType("BODY").setIndex(0)),
        new LayoutPlaceholderIdMapping().setObjectId(rightId).setLayoutPlaceholder(
          new Placeholder().setType("BODY").setIndex(1))
      ).asJava))
    requests ++=
      Seq(
        new Request().setCreateShape(new CreateShapeRequest()
          .setObjectId(topId)
          .setShapeType("TEXT_BOX")
          .setElementProperties(new PageElementProperties()
            .setPageObjectId(slideId)
            .setSize(new Size().setHeight(new Dimension().setMagnitude(3000000).setUnit("EMU")).setWidth(new Dimension().setMagnitude(3000000).setUnit("EMU")))
            .setTransform(new AffineTransform().setScaleX(2.8402).setScaleY(0.3413).setTranslateX(311700).setTranslateY(3545125.0).setUnit("EMU"))
          )
        ),
        new Request().setCreateShape(new CreateShapeRequest()
          .setObjectId(bottomId)
          .setShapeType("TEXT_BOX")
          .setElementProperties(new PageElementProperties()
            .setPageObjectId(slideId)
            .setSize(new Size().setHeight(new Dimension().setMagnitude(3000000).setUnit("EMU")).setWidth(new Dimension().setMagnitude(3000000).setUnit("EMU")))
            .setTransform(new AffineTransform().setScaleX(2.8402).setScaleY(0.3413).setTranslateX(311700).setTranslateY(1152475.0).setUnit("EMU"))
          )
        )
      ) ++ List(leftId, rightId).map(objectId =>
        new Request().setUpdatePageElementTransform(new UpdatePageElementTransformRequest()
          .setObjectId(objectId)
          .setApplyMode("RELATIVE")
          .setTransform(new AffineTransform().setScaleX(1).setScaleY(.4).setTranslateY(1650000).setUnit("EMU"))
        )
      )

    val headings = getHeadings(top)
    assert(headings.size <= 1, "only one heading per slide supported")
    headings.foreach(heading => {
      requests :+= new Request().setInsertText(new InsertTextRequest().setObjectId(titleId).setText(getSlideTitle(heading)))
    })

    requests ++= convertContent(filterHeadings(top), false).genTextAndFormattingRequests(TextContentContext(0, slideId, topId))
    requests ++= convertContent(left, true).genTextAndFormattingRequests(TextContentContext(0, slideId, leftId))
    requests ++= convertContent(right, true).genTextAndFormattingRequests(TextContentContext(0, slideId, rightId))
    requests ++= convertContent(bottom, true).genTextAndFormattingRequests(TextContentContext(0, slideId, bottomId))

    requests ++= createFootnote(references, slideId, List(bottomId))

    (requests, Nil)
  }

  def convertBulletList(bulletList: BulletList, nesting: Int): TextContent = {
    concat(getChildren(bulletList) flatMap {
      case item: ListItem => getChildren(item) map {
        case p: Paragraph => "\t" * nesting +: convertParagraph(p)
        case b: BulletList => convertBulletList(b, nesting + 1)
        case o: OrderedList => convertOrderedList(o, nesting + 1)
      }
      case _ => ???
    })
  }


  def convertOrderedList(bulletList: OrderedList, nesting: Int): TextContent = {
    concat(getChildren(bulletList) flatMap {
      case item: ListItem => getChildren(item) map {
        case p: Paragraph => "\t" * nesting +: convertParagraph(p)
        case b: BulletList => convertBulletList(b, nesting + 1)
        case o: OrderedList => convertOrderedList(o, nesting + 1)
      }
      case _ => ???
    })
  }


  def addImage(url: String, altText: String): TextContent = {
    System.out.println(url)
    def imgId(slideId: String): String = "s"+slideId + "img_" + url.takeRight(15).replace(".", "_").replace("/", "_")
  
    TextContent("").addFormatting(c => List(new Request().setCreateImage(new CreateImageRequest()
      .setObjectId(imgId(c.slideId))
      .setUrl(url)
      .setElementProperties(new PageElementProperties()
        .setPageObjectId(c.slideId)
        //              .setSize(new Size().setHeight(new Dimension().setMagnitude(3000000).setUnit("EMU")).setWidth(new Dimension().setMagnitude(3000000).setUnit("EMU")))
        .setTransform(new AffineTransform().setScaleX(.5).setScaleY(.5).setTranslateX(6000000).setTranslateY(0).setUnit("EMU"))
      )
    ),
      new Request().setUpdatePageElementAltText(new UpdatePageElementAltTextRequest()
        .setObjectId(imgId(c.slideId))
        .setDescription(altText).setTitle(altText)
      )))
  }

  /**
   * top-level call on slide content, do not call recursively
   *
   * handles things that should only occur top level, like images and headings and html directives; everything else is done by convertText */
  private def convertContent(content: Slide, formatSubheadingsBold: Boolean = false): TextContent = {
    val c = content.map {
      case p: Paragraph if p.getFirstChild != null && p.getFirstChild.isInstanceOf[Image] =>
        val image = p.getFirstChild.asInstanceOf[Image]
        val altText = if (image.getFirstChild != null && image.getFirstChild.isInstanceOf[Text])
          image.getFirstChild.asInstanceOf[Text].getLiteral else ""

        val url = image.getDestination
        var img = TextContent(s"[Image: ${url} \"$altText\"]\n").makeTodo()

        val imgURL = if (Set(".png", ".jpeg", ".jpg", ".gif").exists(url.endsWith)) Some(url)
        else if (url endsWith ".webp") Some(url.dropRight(5) + ".png")
        else if (url endsWith ".svg") Some(url.dropRight(4) + ".png")
        else {
          System.err.println(url)
          None
        }
        imgURL foreach { url =>
          img = img concat addImage("https://www.cs.cmu.edu/~ckaestne/fig/01_introduction/" + url, altText)
        }

        img
      case p: Paragraph => convertParagraph(p)
      case bullets: BulletList =>
        convertBulletList(bullets, 0).makeItemize()
      case bullets: OrderedList =>
        convertOrderedList(bullets, 0).makeEnumerate()
      case h: Heading if formatSubheadingsBold =>
        convertInlineText(h.getFirstChild).makeBold() :+ "\n"
      case html: HtmlBlock if matchHtmlComment("discussion")(html) =>
        TextContent("[[Discussion]]\n").makeBold() concat addImage("https://www.cs.cmu.edu/~ckaestne/fig/_assets/discussion.jpg", "Discussion")
      case html: HtmlBlock if Set("col", "colstart", "colend").exists(s => matchHtmlComment(s)(html)) =>
        TextContent("") //ignore
      case html: HtmlBlock if Set("<div class=\"small", "</div", "<!-- .element").exists(html.getLiteral startsWith _) => TextContent("") //ignore divs
      case html: HtmlBlock if html.getLiteral startsWith "<svg" => TextContent("[[SVG not converted, sorry.]]").makeTodo() //ignore divs
      case html: HtmlBlock if html.getLiteral startsWith "<!--" => System.err.println("Ignoring comment: " + html.getLiteral); TextContent("") //ignore divs
      case html: HtmlBlock if html.getLiteral startsWith "<div class=\"tweet\"" => TextContent(s"[[Embedded tweet: ${html.getLiteral.drop(18).dropWhile(_!='"').dropRight(8)}]]").makeTodo() //ignore divs
      case html: HtmlBlock => System.err.println("Unsupported html: "+html.getLiteral); TextContent("")
      case b: BlockQuote => concat(getChildren(b).map(c => convertParagraph(c.asInstanceOf[Paragraph]))).makeItalics().makeIndent()
      case c: FencedCodeBlock =>
        val b = TextContent(c.getLiteral).makeCode()
        if (c.getInfo != null) s"[${c.getInfo}]" +: b
        else b
      //      case allOthers => convertInlineText(allOthers, 0)
    }
    concat(c)
  }

  private def convertParagraph(node: Paragraph /*Paragraph, Emphasis, Text, ...*/): TextContent =
    concat(getChildren(node) map {
      //      case p: Paragraph => concat(getChildren(p).map(c => convertInlineText(c, nesting)))
      //      case bullets: BulletList =>
      //        val items = getChildren(bullets) flatMap {
      //          case item: ListItem => getChildren(item) map {
      //            case p: Paragraph => p
      //          }
      //          case _ => ???
      //        }
      //        concat(items.map(i => convertParagraph(i, nesting + 1)))
      case s: SoftLineBreak => TextContent(" ")
      case i: Image => ???
      case child => convertInlineText(child)
    }) :+ "\n"

  private def convertInlineText(node: Node /*Emphasis, Text, Link...*/): TextContent =
    node match {
      case p: Paragraph => assert(false, "wrong nesting") // ("\t" * nesting +: concat(getChildren(p).map(c => convertInlineText(c, nesting))) :+ "\n")
      //      case bullets: BulletList =>
      //        val items = getChildren(bullets) flatMap {
      //          case item: ListItem => getChildren(item)
      //          case _ => ???
      //        }
      //        concat(items.map(i => convertInlineText(i, nesting + 1)))
      case t: Text => TextContent(t.getLiteral)
      case e: Emphasis => concat(getChildren(e).map(convertInlineText)).makeItalics()
      case e: StrongEmphasis => concat(getChildren(e).map(convertInlineText)).makeBold()
      case l: Link =>
        concat(getChildren(l).map(convertInlineText))
          .setStyle(new TextStyle().setLink(new model.Link().setUrl(l.getDestination)), "link")
      case i: Image => assert(false, s"inline image not supported ${i}")
      case c: Code => TextContent(c.getLiteral).makeCode()
    }


  //  private def convertParagraph(paragraph: Paragraph): TextContent = {
  //    val elements = getChildren(paragraph)
  //    if (elements.size == 1 && elements.head.isInstanceOf[Image]) {
  //
  //    } else convertText(paragraph, 0)
  //  }


  def getChildren(node: Node): List[Node] = if (node == null) Nil else nodeToList(node.getFirstChild)

  def nodeToList(node: Node): List[Node] =
    if (node == null) Nil
    else node :: nodeToList(node.getNext)

  def one[A](l: List[A]): A = {
    if (l.size != 1) throw new IllegalArgumentException("expected list of size 1, but got " + l)
    l.head
  }

  def getSlideTitle(heading: Node): String = {
    getChildren(heading).filter(_.isInstanceOf[Text]).map(_.asInstanceOf[Text].getLiteral).mkString(" ")
  }

  def createFootnote(slidePart: Slide, slideId: String, objectIdsToResize: List[String]): Seq[Request] = if (slidePart.isEmpty) Nil else {
    val content = convertContent(slidePart.tail)
    val footnoteId = slideId + "_footnote"
    Seq(
      new Request().setCreateShape(new CreateShapeRequest()
        .setObjectId(footnoteId)
        .setShapeType("TEXT_BOX")
        .setElementProperties(new PageElementProperties()
          .setPageObjectId(slideId)
          .setSize(new Size().setHeight(new Dimension().setMagnitude(3000000).setUnit("EMU")).setWidth(new Dimension().setMagnitude(3000000).setUnit("EMU")))
          .setTransform(new AffineTransform().setScaleX(2.8402).setScaleY(0.1832).setTranslateX(311700).setTranslateY(4408600.0).setUnit("EMU"))
        )
      )
    ) ++ objectIdsToResize.map(objectId =>
      new Request().setUpdatePageElementTransform(new UpdatePageElementTransformRequest()
        .setObjectId(objectId)
        .setApplyMode("RELATIVE")
        .setTransform(new AffineTransform().setScaleX(1).setScaleY(.95).setTranslateX(0).setTranslateY(0).setUnit("EMU"))
      )
    ) ++ content.genTextAndFormattingRequests(TextContentContext(0,slideId, footnoteId))
  }


  def createNote(slideId: String, content: ConvertMd.TextContent): Presentation => Seq[Request] = (p: Presentation) => if (content.text.isEmpty) Nil else {
    val slide = p.getSlides.asScala.toList.find(_.getObjectId == slideId)
    assert(slide.isDefined, s"slide $slideId not found")
    assert(slide.get.getSlideProperties != null, s"slide $slideId has no properties")
    assert(slide.get.getSlideProperties.getNotesPage != null, s"slide $slideId has no notes page")
    assert(slide.get.getSlideProperties.getNotesPage.getNotesProperties != null, s"slide $slideId has no notes properties")
    assert(slide.get.getSlideProperties.getNotesPage.getNotesProperties.getSpeakerNotesObjectId != null, s"slide $slideId has no notes object id")
    val notesPageId = slide.get.getSlideProperties.getNotesPage.getNotesProperties.getSpeakerNotesObjectId

    content.genTextAndFormattingRequests(TextContentContext(0, slideId,notesPageId))
  }


  def filterHeadings(slide: Slide): Slide = slide.filterNot(_.isInstanceOf[Heading])

}
