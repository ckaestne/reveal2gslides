import com.google.api.client.json.gson.GsonFactory
import com.google.api.services.slides.v1.{Slides, SlidesScopes}
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
import com.google.api.services.drive.{Drive, DriveScopes}

import scala.collection.immutable.List
import com.google.api.services.slides.v1.model.*

import java.io.{File, FileInputStream, FileReader, FileWriter}
import java.util.Collections
import scala.collection.JavaConverters.*
import scala.collection.JavaConverters.*

object GoogleSlidesAPI {
  private val APPLICATION_NAME = "reveal2gslides"
  private val JSON_FACTORY = GsonFactory.getDefaultInstance

  //secret, do not share this file
  private val CREDENTIALS_FILE_PATH = "credentials/api.json"

  lazy val serviceAccountCredentials = ServiceAccountCredentials.fromStream(new FileInputStream(CREDENTIALS_FILE_PATH))

  lazy val credentials: Credentials = serviceAccountCredentials.createScoped(List(SlidesScopes.PRESENTATIONS, DriveScopes.DRIVE).asJava)


  // Create a Google Slides service object
  lazy val slidesService = new Slides.Builder(
    com.google.api.client.googleapis.javanet.GoogleNetHttpTransport.newTrustedTransport(),
    JSON_FACTORY, new HttpCredentialsAdapter(credentials)
  ).setApplicationName("Google Slides Scala Example").build()

  lazy val driveService = new Drive.Builder(
    com.google.api.client.googleapis.javanet.GoogleNetHttpTransport.newTrustedTransport(),
    JSON_FACTORY, new HttpCredentialsAdapter(credentials)
  ).setApplicationName("Drive API Scala Example").build()


  def getSlide(docId: String) =
    // Call the API to retrieve the content of the presentation
    slidesService.presentations().get(docId).execute()

  def saveSlide(docId: String): Unit = {
    val f = new FileWriter("slides.json")
    f.write(getSlide(docId).toString)
    f.close()

  }

  def getSlideIds(docId: String): List[String] = {

    // Call the API to retrieve the content of the presentation
    val presentation = slidesService.presentations().get(docId).execute()
    if (presentation.getSlides == null) return Nil
    presentation.getSlides.asScala.toList.map(_.getObjectId)
  }

  def deleteSlides(docId: String, slideIds: List[String]): Unit = {
    if (slideIds.isEmpty) return
    val deleteSlideRequest = new BatchUpdatePresentationRequest
    deleteSlideRequest.setRequests(slideIds.map(slideId =>
      new Request().setDeleteObject(new DeleteObjectRequest().setObjectId(slideId))
    ).asJava)

    slidesService.presentations().batchUpdate(docId, deleteSlideRequest).execute()
  }


  //  // Define the title and body content for the new slide
  //  //  val titleText = "foo"
  //  //  val bulletPoints = List("x", "y")
  //  //
  //  //  // Create a new slide
  //  //  val slideRequestBody = new Page
  //  //  val slideProperties = new PageProperties
  //  ////  slideProperties.setTitle(titleText)
  //  //  slideRequestBody.setPageProperties(slideProperties)
  //  //  val bulletPointsList = bulletPoints.map { text =>
  //  //    val bullet = new TextContent
  //  //    val bulletParagraph = new Paragraph
  //  //    bulletParagraph.setTextElements(List(
  //  //      new TextElement().setText(text)
  //  //    ).asJava)
  //  //    bullet.setParagraph(bulletParagraph)
  //  //    new List().setNestingLevel(0).setListProperties(
  //  //      new NestingLevel().setBulletAlignment("LEFT")
  //  //    ).setListType("BULLETED").setListId("a").setChildren(List(bullet).asJava)
  //  //  }.asJava
  //  //  val slideElement = new PageElement
  //  //  slideElement.setShape(new Shape)
  //  //  slideElement.getShape.setText(new TextShape)
  //  //  slideElement.getShape.getText.setTextElements(bulletPointsList)
  //  //  slideRequestBody.setPageElements(List(slideElement).asJava)
  //
  def writeSlides(documentId: String, slides: Seq[Request], updates: Seq[Presentation => Seq[Request]] = Nil): Unit = {
    
    
    val batchUpdateResponse = slidesService.presentations().batchUpdate(documentId, new BatchUpdatePresentationRequest().setRequests(slides.asJava)).execute()

    // Get the slide ID of the newly created slide
    //    batchUpdateResponse.getReplies.asScala.foreach(println)


    if (updates.nonEmpty) {
      val presentation = getSlide(documentId)
      val updateSlideRequests = updates.flatMap(_(presentation))
      val batchUpdateResponse2 = slidesService.presentations().batchUpdate(documentId, new BatchUpdatePresentationRequest().setRequests(updateSlideRequests.asJava)).execute()
      batchUpdateResponse2.getReplies.asScala.foreach(println)

    }
  }

  import com.google.api.services.drive.model.Permission

  def createSlide(driveFolder: String, lectureId: String): String = {

    val theme: Presentation = JSON_FACTORY.createJsonParser(new FileReader("theme.json")).parseAndClose(classOf[Presentation])
    val presentation = new Presentation().setTitle(lectureId)./*setPresentationId("mlip" + lectureId).*/setMasters(theme.getMasters)
    val batchUpdateResponse = slidesService.presentations().create(presentation).execute()
    val slideId = batchUpdateResponse.getPresentationId()
    println("https://docs.google.com/presentation/d/"+slideId)
    //set permissions
    driveService.permissions().create(slideId,
      new Permission().setRole("writer").setType("user").setEmailAddress("christian.kaestner@gmail.com")).execute()

    //move to folder
    val fileMetadata = driveService.files().get(slideId).setFields("parents").execute()
    val previousParents = fileMetadata.getParents

    // Update the file's parents to the new folder
    driveService.files().update(slideId, null)
      .setAddParents(driveFolder)
      .setRemoveParents(previousParents.asScala.mkString(","))
      .setFields("id, parents")
      .execute()

    println(s"File $slideId has been moved to folder $driveFolder")

    return slideId


  }

}


