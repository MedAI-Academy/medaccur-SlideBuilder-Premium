plugins {
    id("org.springframework.boot") version "3.2.5"
    id("io.spring.dependency-management") version "1.1.4"
    id("java")
}

group = "com.medaccur"
version = "3.0.0"

java {
    sourceCompatibility = JavaVersion.VERSION_17
    targetCompatibility = JavaVersion.VERSION_17
}

repositories {
    mavenCentral()
}

dependencies {
    // Spring Boot
    implementation("org.springframework.boot:spring-boot-starter-web")
    implementation("org.springframework.boot:spring-boot-starter-actuator")

    // Apache POI 5.3.0 — all three are needed:
    // poi = core, poi-ooxml = XSLF/XSSF classes, poi-ooxml-full = CT* schema beans
    implementation("org.apache.poi:poi:5.3.0")
    implementation("org.apache.poi:poi-ooxml:5.3.0")
    implementation("org.apache.poi:poi-ooxml-full:5.3.0")

    // JSON
    implementation("com.fasterxml.jackson.core:jackson-databind")

    // HTTP client for Chart Microservice calls
    implementation("org.springframework.boot:spring-boot-starter-webflux")

    // Logging
    implementation("org.slf4j:slf4j-api")

    // Test
    testImplementation("org.springframework.boot:spring-boot-starter-test")
}

tasks.withType<Test> {
    useJUnitPlatform()
}

tasks.withType<JavaExec> {
    jvmArgs = listOf("-Xmx512m")
}
