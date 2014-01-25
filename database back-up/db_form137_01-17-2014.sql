-- MySQL dump 10.13  Distrib 5.6.12, for Win32 (x86)
--
-- Host: localhost    Database: db_form137
-- ------------------------------------------------------
-- Server version	5.5.8-log

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `tbl_attendance`
--

DROP TABLE IF EXISTS `tbl_attendance`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_attendance` (
  `SY` varchar(200) NOT NULL,
  `ID` varchar(200) NOT NULL,
  `section_name` varchar(200) NOT NULL,
  `no_school_days` varchar(200) NOT NULL,
  `no_days_absent` varchar(200) NOT NULL,
  `causes_of_absences` varchar(500) NOT NULL,
  `no_days_tardiness` varchar(200) NOT NULL,
  `causes_of_tardiness` varchar(500) NOT NULL,
  `no_days_present` varchar(200) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_attendance`
--

LOCK TABLES `tbl_attendance` WRITE;
/*!40000 ALTER TABLE `tbl_attendance` DISABLE KEYS */;
INSERT INTO `tbl_attendance` VALUES ('2013-2014','109637111111','Apple','180','0','  ','0','  ','180');
/*!40000 ALTER TABLE `tbl_attendance` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_character_grade`
--

DROP TABLE IF EXISTS `tbl_character_grade`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_character_grade` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `SY` varchar(200) NOT NULL,
  `ID` varchar(200) NOT NULL,
  `section_name` varchar(100) NOT NULL,
  `Period` varchar(200) NOT NULL,
  `Honesty` varchar(200) NOT NULL,
  `Courtesy` varchar(200) NOT NULL,
  `Helpfulness_and_Cooperation` varchar(200) NOT NULL,
  `Resourcefulness_and_Creativity` varchar(200) NOT NULL,
  `Consideration_for_Others` varchar(200) NOT NULL,
  `Sportsmanship` varchar(200) NOT NULL,
  `Obedience` varchar(200) NOT NULL,
  `Self_Reliance` varchar(200) NOT NULL,
  `Industry` varchar(200) NOT NULL,
  `Cleanliness_and_Orderliness` varchar(200) NOT NULL,
  `Promptness_and_Punctuality` varchar(200) NOT NULL,
  `Sense_of_Responsibility` varchar(200) NOT NULL,
  `Love_of_God` varchar(200) NOT NULL,
  `Patriotism_and_Love_of_Country` varchar(200) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB AUTO_INCREMENT=11 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_character_grade`
--

LOCK TABLES `tbl_character_grade` WRITE;
/*!40000 ALTER TABLE `tbl_character_grade` DISABLE KEYS */;
INSERT INTO `tbl_character_grade` VALUES (1,'2013-2014','109637111111','Apple','1st Grading','A','B','A','B','A','B','A','B','A','B','A','B','A','B'),(2,'2013-2014','109637222222','Apple','1st Grading','A','A','A','A','A','A','A','A','A','A','A','A','A','A'),(3,'2013-2014','109637111111','Apple','2nd Grading','A','A','A','A','A','A','A','A','A','A','A','A','A','A'),(4,'2013-2014','109637222222','Apple','2nd Grading','B','B','B','B','B','B','B','B','B','B','B','B','B','B'),(5,'2013-2014','109637111111','Apple','Final','A','B','A','B','A','B','A','B','A','B','A','B','A','B'),(6,'2013-2014','109637222222','Apple','Final','','','','','','','','','','','','','',''),(7,'2013-2014','109637111111','Apple','3rd Grading','B','B','B','B','B','B','B','B','B','B','B','B','B','B'),(8,'2013-2014','109637222222','Apple','3rd Grading','','','','','','','','','','','','','',''),(9,'2013-2014','109637111111','Apple','4th Grading','A','A','A','A','A','A','A','A','A','A','A','A','A','A'),(10,'2013-2014','109637222222','Apple','4th Grading','','','','','','','','','','','','','','');
/*!40000 ALTER TABLE `tbl_character_grade` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_grade`
--

DROP TABLE IF EXISTS `tbl_grade`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_grade` (
  `grd_id` int(11) NOT NULL AUTO_INCREMENT,
  `SY` varchar(200) NOT NULL,
  `ID` varchar(200) NOT NULL,
  `section_name` varchar(200) NOT NULL,
  `subject_code` varchar(200) NOT NULL,
  `period` varchar(200) NOT NULL,
  `grade` varchar(100) NOT NULL,
  `remark` varchar(100) NOT NULL,
  PRIMARY KEY (`grd_id`)
) ENGINE=InnoDB AUTO_INCREMENT=26 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_grade`
--

LOCK TABLES `tbl_grade` WRITE;
/*!40000 ALTER TABLE `tbl_grade` DISABLE KEYS */;
INSERT INTO `tbl_grade` VALUES (1,'2013-2014','109637111111','Apple','Eng','1st Grading','90','A'),(2,'2013-2014','109637111111','Apple','Eng','2nd Grading','90','A'),(3,'2013-2014','109637111111','Apple','Eng','3rd Grading','90','A'),(4,'2013-2014','109637111111','Apple','Eng','4th Grading','100','A'),(5,'2013-2014','109637111111','Apple','Eng','Final','92.5','A'),(6,'2013-2014','109637222222','Apple','Eng','1st Grading','90','A'),(7,'2013-2014','109637222222','Apple','Eng','2nd Grading','90','A'),(8,'2013-2014','109637222222','Apple','Eng','3rd Grading','90','A'),(9,'2013-2014','109637222222','Apple','Eng','4th Grading','80','AP'),(10,'2013-2014','109637222222','Apple','Eng','Final','87.5','P'),(11,'2013-2014','109637111111','Apple','Mat','1st Grading','90','A'),(12,'2013-2014','109637111111','Apple','Mat','2nd Grading','90','A'),(13,'2013-2014','109637111111','Apple','Mat','3rd Grading','90','A'),(14,'2013-2014','109637111111','Apple','Mat','4th Grading','85','P'),(15,'2013-2014','109637111111','Apple','Mat','Final','88.75','P'),(16,'2013-2014','109637222222','Apple','Mat','1st Grading','90','A'),(17,'2013-2014','109637222222','Apple','Mat','2nd Grading','90','A'),(18,'2013-2014','109637222222','Apple','Mat','3rd Grading','90','A'),(19,'2013-2014','109637222222','Apple','Mat','4th Grading','95','A'),(20,'2013-2014','109637222222','Apple','Mat','Final','91.25','A'),(21,'2013-2014','109637000006','Cherry','Eng','1st Grading','89','P'),(22,'2013-2014','109637000006','Cherry','Eng','2nd Grading','89','P'),(23,'2013-2014','109637000006','Cherry','Eng','3rd Grading','90','A'),(24,'2013-2014','109637000006','Cherry','Eng','4th Grading','90','A'),(25,'2013-2014','109637000006','Cherry','Eng','Final','89.5','A');
/*!40000 ALTER TABLE `tbl_grade` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_level`
--

DROP TABLE IF EXISTS `tbl_level`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_level` (
  `SY` varchar(100) NOT NULL,
  `lvl_id` int(100) NOT NULL AUTO_INCREMENT,
  `lvl_name` varchar(100) NOT NULL,
  PRIMARY KEY (`lvl_id`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_level`
--

LOCK TABLES `tbl_level` WRITE;
/*!40000 ALTER TABLE `tbl_level` DISABLE KEYS */;
INSERT INTO `tbl_level` VALUES ('2013-2014',1,'Kinder'),('2013-2014',2,'Grade 1'),('2013-2014',3,'Grade 2'),('2014-2015',4,'Grade 1'),('2013-2014',5,'Grade 3'),('2013-2014',6,'Grade 4');
/*!40000 ALTER TABLE `tbl_level` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_logs`
--

DROP TABLE IF EXISTS `tbl_logs`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_logs` (
  `Username` varchar(100) NOT NULL,
  `Login` varchar(100) NOT NULL,
  `Logout` varchar(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_logs`
--

LOCK TABLES `tbl_logs` WRITE;
/*!40000 ALTER TABLE `tbl_logs` DISABLE KEYS */;
INSERT INTO `tbl_logs` VALUES ('admin','1/16/2014 1:22:47 PM','1/16/2014 1:31:21 PM'),('jenny','1/16/2014 1:24:54 PM','1/16/2014 1:50:44 PM'),('ADMIN','1/16/2014 1:31:22 PM','1/16/2014 1:31:54 PM'),('ADMIN','1/16/2014 1:31:54 PM','1/16/2014 1:35:45 PM'),('admin','1/16/2014 1:35:46 PM','1/16/2014 1:37:06 PM'),('admin','1/16/2014 1:39:21 PM','1/16/2014 1:47:25 PM'),('ADMIN','1/16/2014 1:47:25 PM','1/16/2014 3:59:32 PM'),('jenny','1/16/2014 1:50:47 PM','1/16/2014 1:51:21 PM'),('admin','1/16/2014 3:59:34 PM','1/16/2014 4:22:53 PM'),('jenny','1/16/2014 4:08:42 PM','1/17/2014 12:57:52 PM'),('admin','1/16/2014 4:22:53 PM','1/16/2014 4:22:58 PM'),('admin','1/17/2014 12:31:16 PM','1/17/2014 12:32:04 PM'),('ADMIN','1/17/2014 12:32:05 PM','1/17/2014 12:32:11 PM'),('ADMIN','1/17/2014 12:36:42 PM','1/17/2014 12:57:35 PM'),('admin','1/17/2014 12:57:36 PM','1/17/2014 12:58:23 PM'),('jenny','1/17/2014 12:57:53 PM','1/17/2014 1:02:03 PM'),('admin','1/17/2014 12:58:24 PM','1/17/2014 1:17:07 PM'),('jenny','1/17/2014 1:02:04 PM','1/17/2014 1:44:07 PM'),('admin','1/17/2014 1:17:08 PM','1/17/2014 1:45:35 PM'),('jenny','1/17/2014 1:44:08 PM','1/17/2014 3:16:22 PM'),('admin','1/17/2014 1:45:36 PM','1/17/2014 5:23:30 PM'),('jenny','1/17/2014 3:16:22 PM','1/17/2014 4:11:19 PM'),('jenny','1/17/2014 4:11:19 PM','1/17/2014 6:30:31 PM'),('admin','1/17/2014 5:23:32 PM','1/17/2014 5:24:33 PM'),('admin','1/17/2014 6:16:13 PM','1/17/2014 6:33:54 PM'),('jenny','1/17/2014 6:30:32 PM','1/17/2014 6:30:37 PM'),('admin','1/17/2014 6:33:55 PM','1/17/2014 6:33:59 PM'),('admin','1/17/2014 7:02:50 PM','1/17/2014 7:02:55 PM');
/*!40000 ALTER TABLE `tbl_logs` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_promotion`
--

DROP TABLE IF EXISTS `tbl_promotion`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_promotion` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `Name` varchar(100) NOT NULL,
  `ID` varchar(100) NOT NULL,
  `Gender` varchar(100) NOT NULL,
  `SY` varchar(100) NOT NULL,
  `Level` varchar(100) NOT NULL,
  `Section` varchar(100) NOT NULL,
  `Address` varchar(100) NOT NULL,
  `Years_in_School` varchar(100) NOT NULL,
  `Age` varchar(100) NOT NULL,
  `Number_of_Days` varchar(100) NOT NULL,
  `Grade_Remark` varchar(100) NOT NULL,
  `Final_Rating` varchar(100) NOT NULL,
  `Action_Taken` varchar(100) NOT NULL,
  `Remark` varchar(100) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB AUTO_INCREMENT=63 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_promotion`
--

LOCK TABLES `tbl_promotion` WRITE;
/*!40000 ALTER TABLE `tbl_promotion` DISABLE KEYS */;
INSERT INTO `tbl_promotion` VALUES (59,'Sample, Sample Sample','109637333333','Male','2013-2014','Grade 1','Banana','Sample',' 1',' 5.5','0','','No grade','',''),(60,'Camasoza, Charles A','109637111111','Male','2013-2014','Grade 1','Apple','Cavite City',' 1',' 6.75',' 180','A',' 90.62','','A'),(61,'To, Bata a','109637000001','Female','2013-2014','Grade 1','Apple','a',' 1',' 4.75','0','A','No grade','','A'),(62,'Villarde, Markprile John Bbbb','109637222222','Female','2013-2014','Grade 1','Apple','Cavite City',' 1',' 6.25','0','P',' 89.38','','P');
/*!40000 ALTER TABLE `tbl_promotion` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_section`
--

DROP TABLE IF EXISTS `tbl_section`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_section` (
  `SY` varchar(100) NOT NULL,
  `lvl_name` varchar(100) NOT NULL,
  `section_id` int(100) NOT NULL AUTO_INCREMENT,
  `section_name` varchar(100) NOT NULL,
  `teacher_id` varchar(100) NOT NULL,
  PRIMARY KEY (`section_id`)
) ENGINE=InnoDB AUTO_INCREMENT=10 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_section`
--

LOCK TABLES `tbl_section` WRITE;
/*!40000 ALTER TABLE `tbl_section` DISABLE KEYS */;
INSERT INTO `tbl_section` VALUES ('2013-2014','Grade 1',1,'Apple','M-0001'),('2013-2014','Grade 1',2,'Banana','None'),('2013-2014','Grade 1',3,'Cherry','None'),('2013-2014','Grade 2',4,'Sampaguita','None'),('2013-2014','Grade 2',5,'Rose','None'),('2013-2014','Grade 3',6,'Jose Rizal','None'),('2014-2015','Grade 1',7,'Apple','M-0001'),('2013-2014','Kinder',8,'Dog','M-0001'),('2013-2014','Grade 1',9,'Duhat','M-0001');
/*!40000 ALTER TABLE `tbl_section` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_security`
--

DROP TABLE IF EXISTS `tbl_security`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_security` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `Username` varchar(100) NOT NULL,
  `Pet` varchar(100) NOT NULL,
  `Place` varchar(100) NOT NULL,
  `Author` varchar(100) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_security`
--

LOCK TABLES `tbl_security` WRITE;
/*!40000 ALTER TABLE `tbl_security` DISABLE KEYS */;
INSERT INTO `tbl_security` VALUES (1,'admin','aso','bahay','ako'),(2,'M-0001Taneo','a','a','a');
/*!40000 ALTER TABLE `tbl_security` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_student`
--

DROP TABLE IF EXISTS `tbl_student`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_student` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `student_id` varchar(200) NOT NULL,
  `last_name` varchar(200) NOT NULL,
  `first_name` varchar(200) NOT NULL,
  `middle_name` varchar(200) NOT NULL,
  `gender` varchar(200) NOT NULL,
  `bday` varchar(200) NOT NULL,
  `birthplace` varchar(100) NOT NULL,
  `contact_no` varchar(200) NOT NULL,
  `address` varchar(200) NOT NULL,
  `guardian` varchar(200) NOT NULL,
  `guardian_no` varchar(200) NOT NULL,
  `occupation` varchar(100) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_student`
--

LOCK TABLES `tbl_student` WRITE;
/*!40000 ALTER TABLE `tbl_student` DISABLE KEYS */;
INSERT INTO `tbl_student` VALUES (1,'109637111111','Camasoza','Charles','A','Male','2007-03-04','Cavite City','1','Cavite City','A Camasoza','2333','Housewife'),(2,'109637222222','Villarde','Markprile John','Bbbb','Female','2007-08-04','','1','Cavite City','A Villarde','2',''),(4,'109637333333','Sample','Sample','Sample','Male','2008-07-04','','Sample','Sample','Sample','Sample',''),(5,'109637000001','To','Bata','a','Female','2009-02-12','','1','a','a','1',''),(6,'109637000006','Magsaysay','Rexa Acel','','Female','2003-02-17','cavite city','1234567890','Cavite City','Carina Magsaysay','12345','NONE');
/*!40000 ALTER TABLE `tbl_student` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_student_level`
--

DROP TABLE IF EXISTS `tbl_student_level`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_student_level` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `ID` varchar(200) NOT NULL,
  `SY` varchar(200) NOT NULL,
  `lvl_name` varchar(200) NOT NULL,
  `section_name` varchar(200) NOT NULL,
  `Status` varchar(100) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_student_level`
--

LOCK TABLES `tbl_student_level` WRITE;
/*!40000 ALTER TABLE `tbl_student_level` DISABLE KEYS */;
INSERT INTO `tbl_student_level` VALUES (1,'109637111111','2013-2014','Grade 1','Apple','ENROLLED'),(2,'109637222222','2013-2014','Grade 1','Apple','ENROLLED'),(4,'109637333333','2013-2014','Grade 1','Banana','ENROLLED'),(5,'109637000001','2013-2014','Grade 1','Apple','ENROLLED'),(6,'109637000006','2013-2014','Grade 1','Cherry','ENROLLED');
/*!40000 ALTER TABLE `tbl_student_level` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_subject`
--

DROP TABLE IF EXISTS `tbl_subject`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_subject` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `SY` varchar(100) NOT NULL,
  `lvl_name` varchar(100) NOT NULL,
  `subject_code` varchar(100) NOT NULL,
  `subject_name` varchar(100) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_subject`
--

LOCK TABLES `tbl_subject` WRITE;
/*!40000 ALTER TABLE `tbl_subject` DISABLE KEYS */;
INSERT INTO `tbl_subject` VALUES (1,'2013-2014','Grade 1','Eng','English'),(2,'2013-2014','Grade 2','Fil','Filipino'),(4,'2013-2014','Grade 1','Fil','Filipino'),(5,'2013-2014','Grade 1','Mat','Mathematics'),(6,'2013-2014','Kinder','Math','Mathematics');
/*!40000 ALTER TABLE `tbl_subject` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_subjectset`
--

DROP TABLE IF EXISTS `tbl_subjectset`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_subjectset` (
  `No` int(11) NOT NULL AUTO_INCREMENT,
  `SY` varchar(100) NOT NULL,
  `lvl_name` varchar(100) NOT NULL,
  `section_name` varchar(100) NOT NULL,
  `subject_code` varchar(100) NOT NULL,
  `subject_name` varchar(100) NOT NULL,
  `teacher_id` varchar(100) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_subjectset`
--

LOCK TABLES `tbl_subjectset` WRITE;
/*!40000 ALTER TABLE `tbl_subjectset` DISABLE KEYS */;
INSERT INTO `tbl_subjectset` VALUES (1,'2013-2014','Grade 1','Apple','Eng','English','M-0001'),(2,'2013-2014','Grade 1','Apple','Fil','Filipino','M-0001'),(3,'2013-2014','Grade 1','Apple','Mat','Mathematics','M-0002'),(4,'2013-2014','Grade 1','Cherry','Eng','English','M-0001'),(5,'2013-2014','Grade 1','Cherry','Fil','Filipino','M-0002'),(6,'2013-2014','Grade 1','Cherry','Mat','Mathematics','M-0002');
/*!40000 ALTER TABLE `tbl_subjectset` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_sy`
--

DROP TABLE IF EXISTS `tbl_sy`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_sy` (
  `SY` varchar(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_sy`
--

LOCK TABLES `tbl_sy` WRITE;
/*!40000 ALTER TABLE `tbl_sy` DISABLE KEYS */;
INSERT INTO `tbl_sy` VALUES ('2013'),('2014'),('2012'),('2010');
/*!40000 ALTER TABLE `tbl_sy` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_teacher`
--

DROP TABLE IF EXISTS `tbl_teacher`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_teacher` (
  `No` int(11) NOT NULL AUTO_INCREMENT,
  `teacher_id` varchar(100) NOT NULL,
  `first_name` varchar(200) NOT NULL,
  `last_name` varchar(200) NOT NULL,
  `middle_name` varchar(100) NOT NULL,
  `gender` varchar(200) NOT NULL,
  `bday` varchar(200) NOT NULL,
  `contact_no` varchar(200) NOT NULL,
  `course` varchar(200) NOT NULL,
  `school` varchar(200) NOT NULL,
  `a_from` varchar(100) NOT NULL,
  `a_to` varchar(100) NOT NULL,
  `address` varchar(100) NOT NULL,
  `status` varchar(200) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_teacher`
--

LOCK TABLES `tbl_teacher` WRITE;
/*!40000 ALTER TABLE `tbl_teacher` DISABLE KEYS */;
INSERT INTO `tbl_teacher` VALUES (1,'M-0001','Jenny','Taneo','Aaaa','Female','1996-12-04','1111','Bachelor of Science in Information Technology','Cavite State University - Cavite City Campus','2009','2013','Cavite City','On-Duty'),(2,'M-0002','Danie Anne','Pamatian','Bbb','Female','1994-05-04','2222','Bachelor of Scuence in Computer Science','AMA','2000','2004','Cavite City','On-Duty'),(3,'M-0003','Sample','Sample','Aaaabba','Female','2014-01-12','111111111111','','','','','','On-Duty');
/*!40000 ALTER TABLE `tbl_teacher` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_user`
--

DROP TABLE IF EXISTS `tbl_user`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_user` (
  `No` int(100) NOT NULL AUTO_INCREMENT,
  `ID` varchar(100) NOT NULL,
  `Usertype` varchar(100) NOT NULL,
  `Username` varchar(100) NOT NULL,
  `Password` varchar(100) NOT NULL,
  `Lastname` varchar(200) NOT NULL,
  `Firstname` varchar(200) NOT NULL,
  `Middlename` varchar(200) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_user`
--

LOCK TABLES `tbl_user` WRITE;
/*!40000 ALTER TABLE `tbl_user` DISABLE KEYS */;
INSERT INTO `tbl_user` VALUES (1,'1','Administrator','admin','admin','admin','admin','admin'),(2,'M-0001','Teacher','jenny','jenny','Taneo','Jenny','Aaaa'),(3,'M-0002','Teacher','M-0002Pamatian','Pamatian','Pamatian','Danie Anne','Bbb'),(5,'M-0003','Teacher','M-0003Sample','Sample','','','');
/*!40000 ALTER TABLE `tbl_user` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2014-01-17 19:13:51
