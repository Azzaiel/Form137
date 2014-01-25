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
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2014-01-25 18:52:48
