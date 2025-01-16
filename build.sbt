ThisBuild / version := "0.1.0-SNAPSHOT"

ThisBuild / scalaVersion := "3.3.3"

lazy val root = (project in file("."))
  .settings(
    name := "reveal2gslides"
  )

libraryDependencies += "com.google.oauth-client" % "google-oauth-client-jetty" % "1.35.0"
libraryDependencies += "com.google.auth" % "google-auth-library-oauth2-http" % "1.23.0"

libraryDependencies += "com.google.api-client" % "google-api-client" % "1.35.2"

libraryDependencies += "com.google.apis" % "google-api-services-slides" % "v1-rev399-1.25.0"
libraryDependencies += "com.google.apis" % "google-api-services-drive" % "v3-rev197-1.25.0"

libraryDependencies += "org.commonmark" % "commonmark" % "0.21.0"
