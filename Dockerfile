FROM java:8
VOLUME /tmp
ARG JAR_FILE
ADD ${JAR_FILE} word-pdf-service.jar
ADD wait-for-it.sh /wait-for-it.sh
RUN sh -c 'touch /word-pdf-service.jar'
RUN bash -c 'chmod 777 /wait-for-it.sh'
CMD exec java ${JAVA_OPTS}    -Djava.security.egd=file:/dev/./urandom -jar /word-pdf-service.jar