# stream-batch-insert

Reusable Java utilities for:
- streaming CSV generation
- streaming Excel (`.xlsx`) generation
- streaming Excel to JSON conversion

## Maven Dependency

```xml
<dependency>
  <groupId>io.github.prajyotsable</groupId>
  <artifactId>stream-batch-insert</artifactId>
  <version>x.y.z</version>
</dependency>
```

## Version and Deploy

The published dependency version comes from the `<version>` field in `pom.xml`.

Current example:

```xml
<version>1.0.0-SNAPSHOT</version>
```

For a release, change it to a non-snapshot version (for example `1.0.0`) and deploy.

### Option 1: update `pom.xml` manually

1. Update `<version>` in `pom.xml`
2. Run:

```bash
mvn clean deploy
```

### Option 2: update version using Maven

```bash
mvn versions:set -DnewVersion=1.0.0
mvn versions:commit
mvn clean deploy
```

## Sonatype Central Publishing

This project uses:

```xml
<plugin>
  <groupId>org.sonatype.central</groupId>
  <artifactId>central-publishing-maven-plugin</artifactId>
  <version>0.10.0</version>
  <extensions>true</extensions>
  <configuration>
    <publishingServerId>central</publishingServerId>
  </configuration>
</plugin>
```

Make sure your Maven `settings.xml` has credentials for server id `central`.
