if not exists (select * FROM sys.tables where name = N'DOVIZKURLARI' and type = 'U')
   BEGIN
       Create table DOVIZKURLARI (DATE_ DATE,
								  CurrencyCode NVARCHAR(5),
                                  UNIT  VARCHAR(50),
                                  Isim VARCHAR(100),
                                  CurrencyName VARCHAR(100) ,
                                  ForexBuyinh FLOAT  ,
                                  ForexSelling FLOAT,
                                  BanknoteBuying FLOAT,
                                  BanknoteSelling FLOAT)
   end
   GO
  ALTER PROCEDURE[dbo].[DovizKurlari_MerkezBankASi](@TARIH DATE)
   AS
   BEGIN
   WHILE (select COUNT(*) FROM DOVIZKURLARI Where DATE_ = @TARIH) = 0
	BEGIN
	SET @TARIH = DATEADD(DAY,-1,@TARIH)
       DECLARE @url AS VARCHAR(8000)
       DECLARE @XmlYilAy NVARCHAR(6), @XmlTarih NVARCHAR(10)
       SET @XmlYilAy =  CONVERT(nvarchar(6), @TARIH, 112)
       SET @XmlTarih =  FORMAT(@TARIH, 'ddMMyyyy')
       If @TARIH = CONVERT(VARCHAR,GETDATE(),23)
           SET @url =  'https://www.tcmb.gov.tr/kurlar/today.xml'
       ELSE
           SET @url =  'https://www.tcmb.gov.tr/kurlar/' + @XmlYilAy + '/' + @XmlTarih + '.xml'
       DECLARE @OBJ AS INT
       DECLARE @RESULT AS INT
       EXEC @RESULT = SP_OACREATE 'MSXML2.XMLHTTP', @OBJ OUT
       EXEC @RESULT = SP_OAMethod @OBJ , 'open' , null , 'GET', @url, false
       EXEC @RESULT = SP_OAMethod @OBJ, send, NULL,''
       IF OBJECT_ID('tempdb..#XML') IS NOT Null DROP TABLE #XML
       Create table #XML ( STRXML VARCHAR(MAX))
       INSERT INTO #XML(STRXML) EXEC @RESULT = SP_OAGetProperty @OBJ, 'responseXML.xml'
       DECLARE @XML AS XML
       SELECT @XML = STRXML FROM #XML
       DROP TABLE #XML
       DECLARE @HDOC AS INT
       EXEC SP_XML_PREPAREDOCUMENT @HDOC OUTPUT , @XML
       Delete FROM DOVIZKURLARI where DATE_ = @TARIH
       INSERT INTO DOVIZKURLARI ( DATE_,UNIT,Isim,CurrencyName,ForexBuyinh,ForexSelling,BanknoteBuying,BanknoteSelling)
       SELECT @TARIH AS Tarih,* FROM OPENXML(@HDOC, 'Tarih_Date/Currency')
                       With (
                             Unit VARCHAR(50) 'Unit',
                             Isim VARCHAR(100)   'Isim',
                             CurrencyName VARCHAR(100)   'CurrencyName',
                             ForexBuying FLOAT   'ForexBuying',
                             ForexSelling FLOAT 'ForexSelling',
                             BanknoteBuying FLOAT 'BanknoteBuying',
                             BanknoteSelling FLOAT 'BanknoteSelling'
                           )
		 END
   select * FROM DOVIZKURLARI Where DATE_ = (SELECT MAX(DATE_) FROM DOVIZKURLARI WHERE DATE_ <= @TARIH)
	END
GO
EXEC [DovizKurlari_MerkezBankASi] '2020-11-27'
