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
) ENGINE=InnoDB AUTO_INCREMENT=15 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_character_grade`
--

LOCK TABLES `tbl_character_grade` WRITE;
/*!40000 ALTER TABLE `tbl_character_grade` DISABLE KEYS */;
INSERT INTO `tbl_character_grade` VALUES (1,'2013-2014','109637111111','Apple','1st Grading','A','B','A','B','A','B','A','B','A','B','A','B','A','B'),(2,'2013-2014','109637222222','Apple','1st Grading','A','A','A','A','A','A','A','A','A','A','A','A','A','A'),(3,'2013-2014','109637111111','Apple','2nd Grading','A','A','A','A','A','A','A','A','A','A','A','A','A','A'),(4,'2013-2014','109637222222','Apple','2nd Grading','B','B','B','B','B','B','B','B','B','B','B','B','B','B'),(5,'2013-2014','109637111111','Apple','Final','A','B','A','B','A','B','A','B','A','B','A','B','A','B'),(6,'2013-2014','109637222222','Apple','Final','A','C','A','A','A','A','A','A','A','A','A','A','A','A'),(7,'2013-2014','109637111111','Apple','3rd Grading','B','B','B','B','B','B','B','B','B','B','B','B','B','B'),(8,'2013-2014','109637222222','Apple','3rd Grading','','','','','','','','','','','','','',''),(9,'2013-2014','109637111111','Apple','4th Grading','A','A','A','A','A','A','A','A','A','A','A','A','A','A'),(10,'2013-2014','109637222222','Apple','4th Grading','','','','','','','','','','','','','',''),(11,'','109637333333','Banana','1st Grading','A','B','','','','','','','','','','','',''),(12,'','109637333333','Banana','2nd Grading','','','','','','','','','','','','','',''),(13,'','109637333333','Banana','3rd Grading','','','','','','','','','','','','','',''),(14,'','109637333333','Banana','Final','B','B','B','B','B','B','B','B','B','B','B','B','B','B');
/*!40000 ALTER TABLE `tbl_character_grade` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2014-01-25 23:51:44
