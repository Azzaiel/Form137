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
) ENGINE=InnoDB AUTO_INCREMENT=36 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_grade`
--

LOCK TABLES `tbl_grade` WRITE;
/*!40000 ALTER TABLE `tbl_grade` DISABLE KEYS */;
INSERT INTO `tbl_grade` VALUES (1,'2013-2014','109637111111','Apple','Eng','1st Grading','90','A'),(2,'2013-2014','109637111111','Apple','Eng','2nd Grading','90','A'),(3,'2013-2014','109637111111','Apple','Eng','3rd Grading','90','A'),(4,'2013-2014','109637111111','Apple','Eng','4th Grading','100','A'),(5,'2013-2014','109637111111','Apple','Eng','Final','92.5','A'),(6,'2013-2014','109637222222','Apple','Eng','1st Grading','90','A'),(7,'2013-2014','109637222222','Apple','Eng','2nd Grading','90','A'),(8,'2013-2014','109637222222','Apple','Eng','3rd Grading','90','A'),(9,'2013-2014','109637222222','Apple','Eng','4th Grading','80','AP'),(10,'2013-2014','109637222222','Apple','Eng','Final','87.5','P'),(11,'2013-2014','109637111111','Apple','Mat','1st Grading','90','A'),(12,'2013-2014','109637111111','Apple','Mat','2nd Grading','90','A'),(13,'2013-2014','109637111111','Apple','Mat','3rd Grading','90','A'),(14,'2013-2014','109637111111','Apple','Mat','4th Grading','85','P'),(15,'2013-2014','109637111111','Apple','Mat','Final','88.75','P'),(16,'2013-2014','109637222222','Apple','Mat','1st Grading','90','A'),(17,'2013-2014','109637222222','Apple','Mat','2nd Grading','90','A'),(18,'2013-2014','109637222222','Apple','Mat','3rd Grading','90','A'),(19,'2013-2014','109637222222','Apple','Mat','4th Grading','95','A'),(20,'2013-2014','109637222222','Apple','Mat','Final','91.25','A'),(21,'2013-2014','109637000006','Cherry','Eng','1st Grading','89','P'),(22,'2013-2014','109637000006','Cherry','Eng','2nd Grading','89','P'),(23,'2013-2014','109637000006','Cherry','Eng','3rd Grading','90','A'),(24,'2013-2014','109637000006','Cherry','Eng','4th Grading','90','A'),(25,'2013-2014','109637000006','Cherry','Eng','Final','89.5','A'),(26,'','109637111111','Apple','Fil','1st Grading','80','AP'),(27,'','109637111111','Apple','Fil','2nd Grading','80','AP'),(28,'','109637111111','Apple','Fil','3rd Grading','80','AP'),(29,'','109637111111','Apple','Fil','4th Grading','80','AP'),(30,'','109637111111','Apple','Fil','Final','80','AP'),(31,'','109637000001','Apple','Fil','1st Grading','90','A'),(32,'','109637000001','Apple','Fil','2nd Grading','90','A'),(33,'','109637000001','Apple','Fil','3rd Grading','90','A'),(34,'','109637000001','Apple','Fil','4th Grading','90','A'),(35,'','109637000001','Apple','Fil','Final','90','A');
/*!40000 ALTER TABLE `tbl_grade` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2014-01-28 18:48:21
