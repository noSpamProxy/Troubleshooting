SELECT [Thumbprint]
  FROM [NoSpamProxyAddressSynchronization].[CertificateStore].[Certificate]
  WHERE StoreID = '1' and HasPrivateKey = '1' and KeyType = '1'