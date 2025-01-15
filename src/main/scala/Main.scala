import com.google.api.client.json.gson.GsonFactory

object Main extends App {

  val documentId = "1tEX_7XNpHUAFr7M-jJJJod1BUO2g6qFiagioclBWwXU" // ID of the Google Slide presentation you want to access

  import GoogleSlidesAPI._
  import ConvertMd._


  saveSlide(documentId)

  getSlide(documentId)


  val slides = readSlides("/home/ckaestne/Dropbox/work/mlip-s24/lectures/01_introduction/intro.md")

  val slideUpdates = for ((slide, idx) <- slides.drop(2).zipWithIndex) yield {
    convertSlide(slide, s"24_teams_$idx")
  }

//  slideUpdates.flatMap(_._1).zipWithIndex.foreach((e,idx)=> {
//    e.setFactory(GsonFactory.getDefaultInstance)
//    println(idx+": "+e.toPrettyString)
//  })

 deleteSlides(documentId,getSlideIds(documentId))
 writeSlides(documentId, slideUpdates.map(_._1).flatten, slideUpdates.map(_._2).flatten)

  // createSlide("01_introduction")

}
