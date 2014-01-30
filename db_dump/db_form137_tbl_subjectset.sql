CREATE DATABASE  IF NOT EXISTS `db_form137` /*!40100 DEFAULT CHARACTER SET latin1 */;
USE `db_form137`;
-- MySQL dump 10.13  Distrib 5.6.13, for Win32 (x86)
--
-- Host: 127.0.0.1    Database: db_form137
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
INSERT INTO `tbl_subjectset` VALUES (1,'2013-2014','Grade 1','Apple','Eng','English','M-0002'),(2,'2013-2014','Grade 1','Apple','Fil','Filipino','M-0001'),(3,'2013-2014','Grade 1','Apple','Mat','Mathematics','M-0002'),(4,'2013-2014','Grade 1','Cherry','Eng','English','M-0001'),(5,'2013-2014','Grade 1','Cherry','Fil','Filipino','M-0002'),(6,'2013-2014','Grade 1','Cherry','Mat','Mathematics','M-0002');
/*!40000 ALTER TABLE `tbl_subjectset` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2014-01-30 16:23:31
