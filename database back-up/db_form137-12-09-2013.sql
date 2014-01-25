-- MySQL dump 10.13  Distrib 5.6.12, for Win32 (x86)
--
-- Host: localhost    Database: db_form137
-- ------------------------------------------------------
-- Server version	5.6.12-log

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
/*!40000 ALTER TABLE `tbl_attendance` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_character_grade`
--

DROP TABLE IF EXISTS `tbl_character_grade`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_character_grade` (
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
  `Patriotism_and_Love_of_Country` varchar(200) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_character_grade`
--

LOCK TABLES `tbl_character_grade` WRITE;
/*!40000 ALTER TABLE `tbl_character_grade` DISABLE KEYS */;
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
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_grade`
--

LOCK TABLES `tbl_grade` WRITE;
/*!40000 ALTER TABLE `tbl_grade` DISABLE KEYS */;
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
INSERT INTO `tbl_level` VALUES ('2013-2014',4,'Grade 1'),('2013-2014',5,'Grade 2'),('2013-2014',6,'Grade 3');
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
INSERT INTO `tbl_logs` VALUES ('admin','12/4/2013 9:10:32 AM','12/4/2013 9:14:54 AM'),('admin','12/4/2013 9:14:55 AM','12/4/2013 9:16:40 AM'),('admin','12/4/2013 9:16:40 AM','12/4/2013 9:18:14 AM'),('admin','12/4/2013 9:18:15 AM','12/4/2013 9:21:48 AM'),('admin','12/4/2013 9:21:49 AM','12/4/2013 9:43:47 AM'),('admin','12/4/2013 9:43:49 AM','12/4/2013 9:46:40 AM'),('admin','12/4/2013 9:46:40 AM','12/4/2013 10:01:49 AM'),('admin','12/4/2013 10:01:50 AM','12/4/2013 10:02:20 AM'),('admin','12/4/2013 10:02:21 AM','12/4/2013 10:04:30 AM'),('admin','12/4/2013 10:04:31 AM','12/4/2013 10:06:33 AM'),('admin','12/4/2013 10:06:34 AM','12/4/2013 10:09:34 AM'),('admin','12/4/2013 10:09:35 AM','12/4/2013 10:11:05 AM'),('admin','12/4/2013 10:11:06 AM','12/4/2013 10:12:21 AM'),('admin','12/4/2013 10:12:22 AM','12/4/2013 10:14:17 AM'),('admin','12/4/2013 10:14:18 AM','12/4/2013 10:15:27 AM'),('admin','12/4/2013 10:15:28 AM','12/4/2013 10:17:57 AM'),('admin','12/4/2013 10:17:58 AM','12/4/2013 10:20:15 AM'),('admin','12/4/2013 10:20:16 AM','12/4/2013 10:22:15 AM'),('admin','12/4/2013 10:22:16 AM','12/4/2013 10:34:24 AM'),('admin','12/4/2013 10:34:25 AM','12/4/2013 10:35:56 AM'),('admin','12/4/2013 10:35:56 AM','12/4/2013 10:37:08 AM'),('admin','12/4/2013 10:37:08 AM','12/4/2013 9:09:27 PM'),('admin','12/4/2013 9:06:18 PM','12/4/2013 9:09:27 PM'),('admin','12/4/2013 9:09:28 PM','12/4/2013 10:17:51 PM'),('admin','12/4/2013 10:17:52 PM','12/4/2013 11:08:44 PM'),('admin','12/4/2013 11:08:44 PM','12/4/2013 11:12:14 PM'),('admin','12/4/2013 11:12:15 PM','12/4/2013 11:14:52 PM'),('admin','12/4/2013 11:14:53 PM','12/4/2013 11:20:12 PM'),('admin','12/4/2013 11:20:13 PM','12/4/2013 11:30:17 PM'),('admin','12/4/2013 11:30:18 PM','12/4/2013 11:35:41 PM'),('admin','12/4/2013 11:35:42 PM','12/4/2013 11:38:30 PM'),('admin','12/4/2013 11:38:31 PM','None');
/*!40000 ALTER TABLE `tbl_logs` ENABLE KEYS */;
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
INSERT INTO `tbl_section` VALUES ('2013-2014','Grade 1',4,'Apple','None'),('2013-2014','Grade 1',5,'Banana','None'),('2013-2014','Grade 1',6,'Cherry','None'),('2013-2014','Grade 2',7,'Sampaguita','None'),('2013-2014','Grade 2',8,'Rose','None'),('2013-2014','Grade 3',9,'Jose Rizal','None');
/*!40000 ALTER TABLE `tbl_section` ENABLE KEYS */;
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
  `contact_no` varchar(200) NOT NULL,
  `address` varchar(200) NOT NULL,
  `father_name` varchar(200) NOT NULL,
  `father_no` varchar(200) NOT NULL,
  `mother_name` varchar(200) NOT NULL,
  `mother_no` varchar(200) NOT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_student`
--

LOCK TABLES `tbl_student` WRITE;
/*!40000 ALTER TABLE `tbl_student` DISABLE KEYS */;
INSERT INTO `tbl_student` VALUES (1,'S-0001','Camasoza','Charles','A','Male','2007-03-04','1','Cavite City','A Camasoza','2','M Camasoza','2'),(2,'S-0002','Villarde','Markprile John','Bbbb','Female','2007-08-04','1','Cavite City','A Villarde','2','B Villarde','3'),(4,'S-0003','Sample','Sample','Sample','Male','2008-07-04','Sample','Sample','Sample','Sample','Sample','Sample');
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
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_student_level`
--

LOCK TABLES `tbl_student_level` WRITE;
/*!40000 ALTER TABLE `tbl_student_level` DISABLE KEYS */;
INSERT INTO `tbl_student_level` VALUES (1,'S-0001','2013-2014','Grade 1','Apple','ENROLLED'),(2,'S-0002','2013-2014','Grade 1','Apple','ENROLLED'),(4,'S-0003','2013-2014','Grade 1','Banana','ENROLLED');
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
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_subject`
--

LOCK TABLES `tbl_subject` WRITE;
/*!40000 ALTER TABLE `tbl_subject` DISABLE KEYS */;
INSERT INTO `tbl_subject` VALUES (1,'2013-2014','Grade 1','Eng','English'),(2,'2013-2014','Grade 2','Fil','Filipino'),(4,'2013-2014','Grade 1','Fil','Filipino'),(5,'2013-2014','Grade 1','Mat','Mathematics');
/*!40000 ALTER TABLE `tbl_subject` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_subjectset`
--

DROP TABLE IF EXISTS `tbl_subjectset`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_subjectset` (
  `SY` varchar(100) NOT NULL,
  `lvl_name` varchar(100) NOT NULL,
  `section_name` varchar(100) NOT NULL,
  `subject_code` varchar(100) NOT NULL,
  `subject_name` varchar(100) NOT NULL,
  `teacher_id` varchar(100) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_subjectset`
--

LOCK TABLES `tbl_subjectset` WRITE;
/*!40000 ALTER TABLE `tbl_subjectset` DISABLE KEYS */;
INSERT INTO `tbl_subjectset` VALUES ('2013-2014','Grade 1','Apple','Eng','English','None');
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
  `status` varchar(200) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_teacher`
--

LOCK TABLES `tbl_teacher` WRITE;
/*!40000 ALTER TABLE `tbl_teacher` DISABLE KEYS */;
INSERT INTO `tbl_teacher` VALUES ('0001','Jenny','Taneo','Aaaa','Female','1996-12-04','1111','Bachelor of Science in Information Technology','Cavite State University - Cavite City Campus','2009','2013','Cavite City','On-Duty'),('0002','Danie Anne','Pamatian','Bbb','Female','1994-05-04','2222','Bachelor of Scuence in Computer Science','AMA','2000','2004','Cavite City','On-Duty');
/*!40000 ALTER TABLE `tbl_teacher` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `tbl_user`
--

DROP TABLE IF EXISTS `tbl_user`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tbl_user` (
  `ID` varchar(100) NOT NULL,
  `Usertype` varchar(100) NOT NULL,
  `Username` varchar(100) NOT NULL,
  `Password` varchar(100) NOT NULL,
  `Lastname` varchar(200) NOT NULL,
  `Firstname` varchar(200) NOT NULL,
  `Middlename` varchar(200) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_user`
--

LOCK TABLES `tbl_user` WRITE;
/*!40000 ALTER TABLE `tbl_user` DISABLE KEYS */;
INSERT INTO `tbl_user` VALUES ('1','Administrator','admin','admin','admin','admin','admin'),('0001','Teacher','0001Taneo','Taneo','','',''),('0002','Teacher','0002Pamatian','Pamatian','','',''),('2','Administrator','another','a','A','A','Aaa');
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

-- Dump completed on 2013-12-09 22:27:51
