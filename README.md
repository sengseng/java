# java
JxlsUtils.java /**导出大量数据到excel并压缩**/
   1:excel模板文件   jx:area(lastCell="K5")
                     jx:each(items="list" var="item" lastCell="K3")
   2:pom.xml    
   		<dependency>
			<groupId>org.jxls</groupId>
			<artifactId>jxls</artifactId>
			<version>2.3.0</version>
		</dependency>

		<dependency>
			<groupId>org.jxls</groupId>
			<artifactId>jxls-poi</artifactId>
			<version>1.0.9</version>
		</dependency>
    
