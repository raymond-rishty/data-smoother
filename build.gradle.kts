import org.gradle.api.tasks.*
import org.gradle.jvm.tasks.Jar

apply<ApplicationPlugin>()

buildscript {
  repositories {
    gradleScriptKotlin()
    mavenCentral()
  }
  dependencies {
    classpath(kotlinModule("gradle-plugin"))
  }
}

  apply {
    plugin("kotlin")
    plugin<ApplicationPlugin>()
  }

configure<JavaPluginConvention> {
  setSourceCompatibility(1.8)
  setTargetCompatibility(1.8)
}

configure<ApplicationPluginConvention> {
  mainClassName = "datasmoothing.SmoothData"
}

repositories {
  mavenLocal()
  gradleScriptKotlin()
  mavenCentral()
}

dependencies {
  compile("com.google.guava:guava:18.0")
  compile(kotlinModule("stdlib"))
  compile("org.apache.poi:poi:3.14")
  compile("org.apache.poi:poi-ooxml:3.14")
  compile("org.apache.commons:commons-math3:3.6.1")
  compile("com.github.hotchemi:khronos:0.1.0")
  testCompile("junit:junit:4.11")
}


artifacts.add("archives", task<Jar>("sourcesJar") {
  classifier = "sources"
  from(the<JavaPluginConvention>().sourceSets.getByName("main").allSource)
})