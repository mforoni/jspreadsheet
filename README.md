# JSpreadsheet

Provides intuitive Java APIs to manipulate spreadsheets in the ODS and Microsoft Excel formats

## Built With

* [Maven](https://maven.apache.org) - Dependency Management

## Getting Started

### Minimum Requirements

* Java 1.7 or above - tested with [OracleJDK 7.0](http://www.oracle.com/technetwork/java/javase/downloads/java-archive-downloads-javase7-521261.html)
* One build automation tool:
  * [Maven](https://maven.apache.org/download.cgi)
  * [Gradle](https://gradle.org/)

### Adding JSpreadsheet to your build

This project is not yet available on the official maven repository but with [JitPack](https://jitpack.io/) 
it can easily be overcome just by following these two steps:

1. Add the JitPack repository to your build file

```xml
<repositories>
  <repository>
    <id>jitpack.io</id>
    <url>https://jitpack.io</url>
  </repository>
</repositories>
```

1. Add the dependency on JSpreadsheet

```xml
<dependency>
  <groupId>com.github.mforoni</groupId>
  <artifactId>jspreadsheet</artifactId>
  <version>master-SNAPSHOT</version>
</dependency>
```

For Gradle add the following in your root `build.gradle` at the end of repositories:

```gradle
allprojects {
  repositories {
    ...
    maven { url 'https://jitpack.io' }
  }
}

dependencies {
  implementation 'com.github.mforoni:jspreadsheet:master-SNAPSHOT'
}
```

## Code Style

This project follow the [Google Java Style Guide](https://google.github.io/styleguide/javaguide.html).

## Project Management

[Trello Board.](https://trello.com/b/rdOZh4ci/jspreadsheet)

## Author

* **Marco Foroni** - [mforoni](https://github.com/mforoni)

## License

This project is licensed under the MIT License - see the [LICENSE.md](https://github.com/mforoni/jspreadsheet/blob/master/LICENSE) file for details


## IMPORTANT WARNINGS

1. This project is under development.

1. APIs marked with the `@Beta` annotation at the class or method level
are subject to change. They can be modified in any way, or even
removed, at any time. Read more about [`@Beta`](https://github.com/google/guava#important-warnings) annotation.
