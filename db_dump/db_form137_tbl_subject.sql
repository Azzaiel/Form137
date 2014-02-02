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
  `last_mod_date` datetime DEFAULT NULL,
  PRIMARY KEY (`No`)
) ENGINE=InnoDB AUTO_INCREMENT=89 DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `tbl_subject`
--

LOCK TABLES `tbl_subject` WRITE;
/*!40000 ALTER TABLE `tbl_subject` DISABLE KEYS */;
INSERT INTO `tbl_subject` VALUES (9,'','Grade 1','Math','Math','2014-01-30 12:50:15'),(10,'','Kinder','Cognitive Development','Cognitive Development','2014-01-30 12:50:15'),(11,'','Kinder','Pyschomotor Development','Pyschomotor Development','2014-01-30 12:50:15'),(12,'','Kinder','Social and Emotional Development','Social and Emotional Development','2014-01-30 12:50:15'),(13,'','Grade 1','Filipino','Filipino','2014-01-30 12:50:15'),(14,'','Grade 1','English','English','2014-01-30 12:50:15'),(15,'','Grade 1','A.P','A.P','2014-01-30 12:50:15'),(16,'','Grade 1','Mapeh','Mapeh','2014-01-30 12:50:15'),(17,'','Grade 1','Arts','Arts','2014-01-30 12:50:15'),(18,'','Grade 1','P.E','P.E','2014-01-30 12:50:15'),(19,'','Grade 1','Health','Health','2014-01-30 12:50:15'),(20,'','Grade 1','Edukasyon sa Pagpapakatao','Edukasyon sa Pagpapakatao','2014-01-30 12:50:15'),(22,'','Grade 2','Math','Math','2014-01-30 12:50:15'),(23,'','Grade 2','Filipino','Filipino','2014-01-30 12:50:15'),(24,'','Grade 2','English','English','2014-01-30 12:50:15'),(25,'','Grade 2','A.P','A.P','2014-01-30 12:50:15'),(26,'','Grade 2','Mapeh','Mapeh','2014-01-30 12:50:15'),(27,'','Grade 2','Arts','Arts','2014-01-30 12:50:15'),(28,'','Grade 2','P.E','P.E','2014-01-30 12:50:15'),(29,'','Grade 2','Health','Health','2014-01-30 12:50:15'),(30,'','Grade 2','Edukasyon sa Pagpapakatao','Edukasyon sa Pagpapakatao','2014-01-30 12:50:15'),(34,'','Grade 3','English','English','2014-01-30 12:50:15'),(35,'','Grade 3','A.P','A.P','2014-01-30 12:50:15'),(36,'','Grade 3','Mapeh','Mapeh','2014-01-30 12:50:15'),(37,'','Grade 3','Arts','Arts','2014-01-30 12:50:15'),(38,'','Grade 3','P.E','P.E','2014-01-30 12:50:15'),(39,'','Grade 3','Health','Health','2014-01-30 12:50:15'),(40,'','Grade 3','Edukasyon sa Pagpapakatao','Edukasyon sa Pagpapakatao','2014-01-30 12:50:15'),(60,'','Grade 4','Math','Math','2014-01-30 12:50:15'),(61,'','Grade 4','Filipino','Filipino','2014-01-30 12:50:15'),(62,'','Grade 4','English','English','2014-01-30 12:50:15'),(63,'','Grade 4','A.P','A.P','2014-01-30 12:50:15'),(64,'','Grade 4','Mapeh','Mapeh','2014-01-30 12:50:15'),(65,'','Grade 4','Arts','Arts','2014-01-30 12:50:15'),(66,'','Grade 4','P.E','P.E','2014-01-30 12:50:15'),(67,'','Grade 4','Health','Health','2014-01-30 12:50:15'),(68,'','Grade 4','Edukasyon sa Pagpapakatao','Edukasyon sa Pagpapakatao','2014-01-30 12:50:15'),(70,'','Grade 5','Math','Math','2014-01-30 12:50:15'),(71,'','Grade 5','Filipino','Filipino','2014-01-30 12:50:15'),(72,'','Grade 5','English','English','2014-01-30 12:50:15'),(73,'','Grade 5','A.P','A.P','2014-01-30 12:50:15'),(74,'','Grade 5','Mapeh','Mapeh','2014-01-30 12:50:15'),(75,'','Grade 5','Arts','Arts','2014-01-30 12:50:15'),(76,'','Grade 5','P.E','P.E','2014-01-30 12:50:15'),(77,'','Grade 5','Health','Health','2014-01-30 12:50:15'),(78,'','Grade 5','Edukasyon sa Pagpapakatao','Edukasyon sa Pagpapakatao','2014-01-30 12:50:15'),(80,'','Grade 6','Math','Math','2014-01-30 12:50:15'),(81,'','Grade 6','Filipino','Filipino','2014-01-30 12:50:15'),(82,'','Grade 6','English','English','2014-01-30 12:50:15'),(83,'','Grade 6','A.P','A.P','2014-01-30 12:50:15'),(84,'','Grade 6','Mapeh','Mapeh','2014-01-30 12:50:15'),(85,'','Grade 6','Arts','Arts','2014-01-30 12:50:15'),(86,'','Grade 6','P.E','P.E','2014-01-30 12:50:15'),(87,'','Grade 6','Health','Health','2014-01-30 12:50:15'),(88,'','Grade 6','Edukasyon sa Pagpapakatao','Edukasyon sa Pagpapakatao','2014-01-30 12:50:15');
/*!40000 ALTER TABLE `tbl_subject` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2014-02-02 10:17:58
