import GoogleSlidesAPI.{createSlide, deleteSlides, getSlideIds, writeSlides}
import com.google.api.client.json.gson.GsonFactory

import java.io.FileWriter

object Main extends App {

  val driveFolder = "1vv0TfPlfe8-bAv5DehNbxBxlR2bmWKeX"

  val lectures =
    """
      |./01_introduction/intro.md
      |./02_systems/systems.md
      |./03_requirements/requirements.md
      |./04_mistakes/mistakes.md
      |./05_modelaccuracy/modelquality1.md
      |./06_teamwork/teams.md
      |./07_modeltesting/modelquality2.md
      |./08_architecture/tradeoffs.md
      |./09_deploying_a_model/deployment.md
      |./10_qainproduction/qainproduction.md
      |./11_dataquality/dataquality.md
      |./12_pipelinequality/pipelinequality.md
      |./13_dataatscale/dataatscale.md
      |./14_operations/operations.md
      |./15_provenance/provenance.md
      |./16_process/process.md
      |./17_intro_ethics_fairness/intro-ethics-fairness.md
      |./18_fairness_measures/model_fairness.md
      |./19_system_fairness/system_fairness.md
      |./20_explainability/explainability.md
      |./21_transparency/transparency.md
      |./22_security/security.md
      |./23_safety/safety.md
      |./24_summary/all.md
      |./24_teams/teams.md
      |""".stripMargin.split("\n").filter(_.nonEmpty).map(_.drop(2))


  for (lecture <- lectures.drop(24).take(8)) {
    val lectureId = lecture.split("/").head
    System.out.println("### " + lectureId)

    val converter = new ConvertMd(lectureId, s"https://www.cs.cmu.edu/~ckaestne/fig/$lectureId/")

    val slides = converter.readSlides("/home/ckaestne/Dropbox/work/mlip/lectures/" + lecture)
    val slideUpdates = for ((slide, idx) <- slides.drop(2).zipWithIndex)
      yield converter.convertSlide(slide, idx)

    val updateCommands = new StringBuffer()
    slideUpdates.flatMap(_._1).zipWithIndex.foreach((e, idx) => {
      e.setFactory(GsonFactory.getDefaultInstance)
      updateCommands.append(idx).append(": ").append(e.toPrettyString)
    })
    // write updateCommands to a file
    val f = new FileWriter(s"updateCommands_${lectureId}.tmp")
    f.write(updateCommands.toString)
    f.close()

    val documentId = createSlide(driveFolder, lectureId)
    //    deleteSlides(documentId, getSlideIds(documentId))
    writeSlides(documentId, slideUpdates.map(_._1).flatten, slideUpdates.map(_._2).flatten)
  }


//    import GoogleSlidesAPI._
//      val documentId = "1tEX_7XNpHUAFr7M-jJJJod1BUO2g6qFiagioclBWwXU" // ID of the Google Slide presentation you want to access
//
//  ////  saveSlide(documentId)
//  //
//  ////  getSlide(documentId)
//  //
//    val converter = new ConvertMd("05_modelaccuracy", "https://www.cs.cmu.edu/~ckaestne/fig/05_modelaccuracy/")
//
//    val slides = converter.readSlides("/home/ckaestne/Dropbox/work/mlip/lectures/05_modelaccuracy/modelquality1.md")
//
//    val slideUpdates = for ((slide, idx) <- slides.drop(21).take(1).zipWithIndex)
//      yield converter.convertSlide(slide, idx)
//
//    val updateCommands = new StringBuffer()
//    slideUpdates.flatMap(_._1).zipWithIndex.foreach((e,idx)=> {
//      e.setFactory(GsonFactory.getDefaultInstance)
//      updateCommands.append(idx).append(": ").append(e.toPrettyString)
//    })
//    // write updateCommands to a file
//    val f = new FileWriter(s"updateCommands.tmp")
//    f.write(updateCommands.toString)
//    f.close()
//  //
//   deleteSlides(documentId,getSlideIds(documentId))
//   writeSlides(documentId, slideUpdates.map(_._1).flatten, slideUpdates.map(_._2).flatten)
//   System.out.println(s"Slides ${converter.slideDeckId} published to https://docs.google.com/presentation/d/$documentId")
//  //
//  //
//  //   createSlide(driveFolder, "01_introduction")

}
