
'***************************************
' Создать заказ по заданным параметрам
'  Входные Параметры:
'    rstMeterReq - Набор Данных заявки на замер
'    lngUserID - ID пользователя, от имени которого создаётся заказ
'    lngProfileID - ID профиля, для которого создаётся заказ
'    lngTradeAgentID - ID торгового представителя, для начисления бонусов
'    strDCardNum -
' Выходные параметры:
' Возвращает № созданного заказа или описание ошибки
'***************************************
Public Function CreateOrderByMetReq(ByVal rstMeterReq As Recordset, _
                                    Optional ByVal lngUserID As Long = 0, _
                                    Optional ByVal lngProfileID As Long = 0, _
                                    Optional ByVal lngTradeAgentID As Long = 0, _
                                    Optional ByVal strDCardNum As String = "") As Long
  Dim strReturnString As String
  Dim lngOrderID As Long
  Dim lngPayCondID As Long
  Dim dblExchange As Double
  Dim strCommandNumber As String
  Dim lngCurrencyID As Long
  Dim strMetReqNum As String
  Dim lngMeterID As Long
  Dim lngDocFormID As Long
  Dim lngManagerID As Long
  Dim lngOfficeID As Long
  Dim lngDepartID As Long
  Dim lngSellerID As Long
  Dim lngCarrierID As Long
  Dim lngMounterID As Long
  Dim lngMakerID As Long
  Dim lngWinMakerID As Long
  Dim lngGlassMakerID As Long
  Dim lngAEMakerID As Long
  Dim lngTradeMarkID As Long
  Dim rc As Long
  Dim rstOrder As Recordset
  Dim rstProps As Recordset
  Dim rstProfile As Recordset
  Dim objOrdMileSt As clsOrdMileSt
  Dim rstOrdMilest As Recordset
  Dim objCurrency As clsCCurrency
  Dim objCRate As clsCRate
  Dim objCodif  As clsCodificator
  Dim rstCRate As Recordset
  Dim objDealerData  As clsDealer
  Dim rstDealerData  As Recordset
  Dim objMetReq As clsMetering
  Dim colMSDates As Collection
  Dim rstVendor As Recordset
  Dim rstMaker As Recordset
  Dim objUser As clsUser
  Dim objProp As clsOrderProp
  Dim objFcdOrder As fcdOrder
  Dim lngMoveConstrRuleID As Long
  Dim lngStockID As Long
  Dim lngShipWayID As Long
  Dim lngContrTypeID As Long
  Dim lngContrTypeID1 As Long
  Set objProp = New clsOrderProp
  Set objOrdMileSt = New clsOrdMileSt
  objOrdMileSt.BCData = objBCData
  Set objCRate = New clsCRate
  Set objCodif = New clsCodificator
  Set objUser = New clsUser
  Set objFcdOrder = New fcdOrder
  strReturnString = ""
  ' передать ссылку на общие данные
  objCRate.BCData = objBCData
  objCodif.BCData = objBCData
  objUser.BCData = objBCData
  objProp.BCData = objBCData
  objFcdOrder.BCData = objBCData
  ' получить курсы на текущую дату
  Set rstCRate = objCRate.LoadByDateInInterval(Date)
  strCommandNumber = Trim(rstCRate!strCommandNumber & vbNullString)
  dblExchange = objCRate.GetRateByDate(Date, TRATE_EUR)
  ' условия оплаты = "по стандартной скидке"
  lngPayCondID = DBProcLib.GetIDByKey(objBCData, "PAYMENTCOND", "STDSC")
  ' если есть наряд по заказу - подтянуть оттуда данные исполнителя и по ним получить пользователя!!
  If lngUserID = 0 Then lngUserID = objBCData.UserSID
  If lngProfileID <> 0 Then
    Set rstProfile = objUser.LoadP(lngProfileID)
  Else
    Set rstProfile = objUser.LoadProfileByUserID(lngUserID, False, True)
  End If
  lngMoveConstrRuleID = 0
  lngStockID = 0
  lngTradeMarkID = 0
  lngContrTypeID = 0
  If rstProfile.RecordCount > 0 Then
    lngManagerID = IIf(IsNull(rstProfile!lngManagerID), 0, rstProfile!lngManagerID)
    lngOfficeID = IIf(IsNull(rstProfile!lngOfficeID), 0, rstProfile!lngOfficeID)
    lngDepartID = IIf(IsNull(rstProfile!lngDepartID), 0, rstProfile!lngDepartID)
    lngSellerID = IIf(IsNull(rstProfile!lngSellerID), 0, rstProfile!lngSellerID)
    lngCarrierID = IIf(IsNull(rstProfile!lngCarrierID), 0, rstProfile!lngCarrierID)
    lngMounterID = IIf(IsNull(rstProfile!lngMounterID), 0, rstProfile!lngMounterID)
    lngMakerID = IIf(IsNull(rstProfile!lngMakerID), 0, rstProfile!lngMakerID)
    lngWinMakerID = IIf(IsNull(rstProfile!lngWinMakerID), 0, rstProfile!lngWinMakerID)
    lngGlassMakerID = IIf(IsNull(rstProfile!lngGlassMakerID), 0, rstProfile!lngGlassMakerID)
    lngAEMakerID = IIf(IsNull(rstProfile!lngAEMakerID), 0, rstProfile!lngAEMakerID)
    lngMoveConstrRuleID = IIf(IsNull(rstProfile!lngMoveConstrRuleID), 0, rstProfile!lngMoveConstrRuleID)
    lngStockID = IIf(IsNull(rstProfile!lngStockID), 0, rstProfile!lngStockID)
    lngTradeMarkID = IIf(IsNull(rstProfile!lngTradeMarkID), 0, rstProfile!lngTradeMarkID)
    lngContrTypeID = IIf(IsNull(rstProfile!lngContrTypeID), 0, rstProfile!lngContrTypeID)
  End If
  
  If lngManagerID = 0 Then lngManagerID = objBCData.ManagersID
  If lngOfficeID = 0 Then lngOfficeID = objBCData.OfficeSID
  If lngDepartID = 0 Then lngDepartID = objBCData.DepartmentSID
  If lngSellerID = 0 Then lngSellerID = objBCData.SellerSID
  If lngCarrierID = 0 Then lngCarrierID = objBCData.CarrierSID
  If lngMounterID = 0 Then lngMounterID = objBCData.MounterSID
  If lngMakerID = 0 Then lngMakerID = objBCData.MakerSID
  lngCurrencyID = objCodif.GetIDByKey("CURRENCY", TRATE_RUR)

  Dim lngOtmp As Long
  Dim strOnum As String
  Dim strOnumA As String
  Dim fIsChangeDF As Boolean
  fIsChangeDF = False
  lngOtmp = objBCData.OfficeSID
  strOnum = objBCData.OfficeNum
  strOnumA = objBCData.OfficeNumAlias
  objBCData.OfficeSID = lngOfficeID
  If rstProfile.RecordCount > 0 Then
    objBCData.OfficeNum = Trim(rstProfile!strPrefNumber & vbNullString)
    objBCData.OfficeNumAlias = Trim(rstProfile!strPrefNumberAls & vbNullString)
  End If
  
  Set rstOrder = Me.Load(0)
  objBCData.OfficeSID = lngOtmp
  objBCData.OfficeNum = strOnum
  objBCData.OfficeNumAlias = strOnumA
  
  ' задать торгового агента
  If lngTradeAgentID <> 0 Then
    rstOrder!lngTradeAgentID = lngTradeAgentID
    Dim objTA As New clsTradeAgents
    objTA.BCData = objBCData
    rstOrder!strAgentCode = objTA.GetCodeByID(lngTradeAgentID)
    Set objTA = Nothing
  End If
  ' задать номер д.карты
  If Trim(strDCardNum) <> "" Then
    rstOrder!strDiscountCard = strDCardNum
  End If
  
  'если заказ по замеру платного ремонта и и менеджер заказа принадлежит к группе Замерщики
  'то в договор прописать офис "Платный ремонт"
  If Trim(rstMeterReq!strContactKindKey & vbNullString) = "RP" And objUser.GetUserGroupKeyByManagerID(lngManagerID) = "METER" And _
      rstMeterReq!lngTradeMarkID = DBProcLib.GetIDByKey(objBCData, "TRADEMARK", "MO") Then
    lngOfficeID = 125
  End If
  Dim strComment As String
  Dim n1 As Long
    Set rstProps = objProp.Load(0)
    rstOrder!datOrderDate = Date
    rstOrder!txtContact = rstMeterReq!strContact
    rstOrder!txtPostAddr = rstMeterReq!strPost
    rstOrder!strEmail = rstMeterReq!strEmail
    rstOrder!txtPhone = rstMeterReq!strPhone
    rstOrder!strPhone1 = rstMeterReq!strPhone1
    rstOrder!strPhone2 = rstMeterReq!strPhone2
    rstOrder!lngKladdrID = rstMeterReq!lngKladdrID
    rstOrder!lngRegionID = rstMeterReq!lngRegionID
    strComment = Trim(rstMeterReq!strComment & vbNullString)
    n1 = InStr(1, strComment, "Отправитель: visitor")
    If n1 > 0 Then
      rstOrder!txtComment = Left(strComment, n1 - 1) '& " " & Mid(strComment, n2 + Len("</#COMMON>"))
    Else
      rstOrder!txtComment = strComment
    End If
    'rstOrder!txtComment = rstMeterReq!strComment
    rstOrder!intProdQuant = 0 'rstMeterReq!intQuantity
    rstOrder!lngManagerID = lngManagerID  ' rstMeterReq!lngManagerID
    rstOrder!lngOperatorID = lngManagerID
    rstOrder!lngOfficeID = lngOfficeID 'rstMeterReq!lngOfficeID
    rstOrder!lngDepartID = lngDepartID 'rstMeterReq!lngDepartID
    rstOrder!lngAdvertSrcID = rstMeterReq!lngAdvertSrcID
    If rstProfile.RecordCount > 0 Then
      rstOrder!lngDlvrMethodID = rstProfile!lngDlvrMethodID
    End If
    If RSProcLib.IsField(rstMeterReq, "lngOrgID") Then
      If Not IsNull(rstMeterReq!lngOrgID) Then
        rstOrder!lngOrgID = rstMeterReq!lngOrgID
      End If
    End If
    If RSProcLib.IsField(rstMeterReq, "lngTradeMarkID") Then
      If Not IsNull(rstMeterReq!lngTradeMarkID) Then
        lngTradeMarkID = rstMeterReq!lngTradeMarkID
      End If
    End If
    DBProcLib.GetDefaultByMakerID objBCData, lngMakerID, lngStockID, lngShipWayID, 0, 0, 0, 0, 0
    rstOrder!lngMoveConstrRuleID = IIf(lngMoveConstrRuleID = 0, Null, lngMoveConstrRuleID)
    rstOrder!lngStockID = IIf(lngStockID = 0, Null, lngStockID)
    rstOrder!lngShipWayID = IIf(lngShipWayID = 0, Null, lngShipWayID)
    rstOrder!lngTradeMarkID = IIf(lngTradeMarkID = 0, Null, lngTradeMarkID)
    rstOrder!boolIsBookReg = False

    rstOrder!curExchange = dblExchange
    If lngPayCondID <> 0 Then
      rstOrder!lngPayCondID = lngPayCondID
    End If
    'валюта заказа
    If lngCurrencyID <> 0 Then
      rstOrder!lngCurrencyID = lngCurrencyID
    End If
    rstOrder!strCommandNumber = strCommandNumber
    rstOrder!lngCurMileStone = objOrdMileSt.lngOrderMSID
	
    ' тип договора брать из настроек менеджера
    ' задать в заказе тип договора - Обычный договор
    lngContrTypeID1 = objFcdOrder.FindContrType(rstMeterReq, "", True, lngManagerID)
    If lngContrTypeID1 > 0 Then
      rstOrder!lngContrTypeID = lngContrTypeID1
    Else
      rstOrder!lngContrTypeID = IIf(lngContrTypeID = 0, Null, lngContrTypeID)  'ORD_SUPPLYAGREEMENT
    End If
    ' переписать признак 100% переноса доставки из формы документа
    rstOrder!boolIs100Prc = DBProcLib.Get100PrcByDocFormID(objBCData, lngDocFormID)
    ' номер замера
    strMetReqNum = Trim(rstMeterReq!strMetReqNum & vbNullString)
    ' прописать наименование заказа
    rstOrder!txtName = Left(Trim(rstMeterReq!strName), rstOrder.Fields("txtName").DefinedSize)
    ' взять данные из текущего офиса
    rstOrder!lngSellerID = lngSellerID
    rstOrder!lngCarrierID = lngCarrierID
    rstOrder!lngMounterID = lngMounterID
    rstOrder!lngMakerID = lngMakerID
     
        DBProcLib.GetDefaultByMakerID objBCData, lngMakerID, lngStockID, lngShipWayID, 0, 0, 0, 0, 0
        If lngShipWayID = 0 Then
        ' получить ID пункта доставки Москва и Московская Область
          lngShipWayID = DBProcLib.GetIDByKey(objBCData, "SHIPWAY", "1MO")
        End If
        ' если ID доставки получен
        If lngShipWayID <> 0 Then
          ' записать в заказ ID пункта доставки по-умолчанию
          rstOrder!lngShipWayID = lngShipWayID
        End If
        'lngStockID = DBProcLib.GetDefaultStockID(objBCData, objBCData.ManagersID, rstRecord!lngTradeMarkID)
        ' если ID доставки получен
        If lngStockID <> 0 Then
          ' записать в заказ ID пункта доставки по-умолчанию
          rstOrder!lngStockID = lngStockID
        End If
    ' получить данные продавца и изготовителя
    Set rstVendor = DBProcLib.GetVendorData(objBCData, lngSellerID)
    Set rstMaker = DBProcLib.GetVendorData(objBCData, lngMakerID)

  
    rstOrder!dblProdCoef = IIf(objBCData.GetConstantValue("UseCostST"), 0, rstMaker!dblSpecFactor)
    rstOrder!dblSaleCoef = rstVendor!dblSpecFactor
    rstOrder!curProdCostR = 0
    rstOrder!curTax = 0
    rstOrder!curProdTotCost = 0
    rstOrder!curServCostR = 0 'curMeterCost
    rstOrder!curServCost = 0 ' Round(curMeterCost / dblExchange, 2)
    rstOrder!curOrdTotCost = rstOrder!curProdTotCost
    rstOrder!curMountCostR = 0
    rstOrder!curMountCost = 0
    rstOrder!curDlvrCostR = 0
    rstOrder!curDlvrCost = 0
    rstOrder!sngSquare = 0
    rstOrder!intPosition = 0
    rstOrder!intProdQuant = 0
    ' номер заявки на замер = введенный номер,
    rstOrder!strMetReqNum = rstMeterReq!strMetReqNum
    If IsCanChangeDocForm(rstMeterReq!lngRecordID, rstOrder!lngContrTypeID) Then
      Dim objTradeMarkProp As clsTradeMarkProp
      Dim objDocForm As clsDocForm
      Set objTradeMarkProp = New clsTradeMarkProp
      objTradeMarkProp.BCData = objBCData
     Set objDocForm = New clsDocForm
      objDocForm.BCData = objBCData
      fIsChangeDF = True
      rstOrder!lngDocFormID = DBProcLib.GetIDByKey(objBCData, "DOCFORM", "OINVAG")
      SetAgentMounterByDocForm rstProps, rstOrder
	  
      Set objTradeMarkProp = Nothing
      Set objDocForm = Nothing

    End If
    ' записать заказ в БД, получить его ID
    lngOrderID = objOrder.UpLoad(rstOrder)
    ' если заказ не сохранился
    If lngOrderID = -1 Then
      Exit Function
    End If
    
    If Not fIsChangeDF Then rstProps!lngSellerID = lngSellerID
    rstProps!lngCarrierID = lngCarrierID
    If Not fIsChangeDF Then rstProps!lngMounterID = lngMounterID
    rstProps!lngMakerID = lngMakerID
    If rstMaker.RecordCount > 0 And lngWinMakerID = 0 Then
      If Not IsNull(rstMaker!lngLinkID) Then
        lngWinMakerID = rstMaker!lngLinkID
      End If
    End If
    rstProps!lngWinMakerID = lngWinMakerID
    rstProps!lngGlassMakerID = lngGlassMakerID
    rstProps!lngAEMakerID = lngAEMakerID
    ' установить привязку атрибутов к заказу
    rstProps!lngOrderID = lngOrderID
    rc = objProp.UpLoad(rstProps)

    ' создать связь с заказа замером
    objOrder.SetLinkToMetReq rstMeterReq!lngRecordID, lngOrderID
    
    '-технологические этапы = оформление, изготовление, комплектование,[транспортировка,] доставка
    'даты этапов рассчитываются согласно стандартной длительности (из справочника тех.этапов), начиная с текущей даты.
    Set colMSDates = New Collection
    RecalcMSDate Date, colMSDates
    ' Добавить этапы
    ' Оформление
    Set rstOrdMilest = objOrdMileSt.LoadUPD(0)
    rstOrdMilest!lngOrderID = lngOrderID
    
    rstOrdMilest!datActDate = colMSDates.Item(1)
    rstOrdMilest!datPlanDate = rstOrdMilest!datActDate
    rstOrdMilest!lngTechMilestID = objOrdMileSt.lngOrderMSID
    rstOrdMilest!boolIsFinish = False
    rstOrdMilest!intSequenceNum = objOrdMileSt.GetTMSequenceByID(rstOrdMilest!lngTechMilestID)
    rc = objOrdMileSt.UpLoad(rstOrdMilest)
    ' Изготовление
    Set rstOrdMilest = objOrdMileSt.LoadUPD(0)
    rstOrdMilest!lngOrderID = lngOrderID
    rstOrdMilest!datActDate = colMSDates.Item(2)
    rstOrdMilest!datPlanDate = rstOrdMilest!datActDate
    rstOrdMilest!lngTechMilestID = objOrdMileSt.lngConstrMSID
    rstOrdMilest!intSequenceNum = objOrdMileSt.GetTMSequenceByID(rstOrdMilest!lngTechMilestID)
    rstOrdMilest!boolIsFinish = False
    rc = objOrdMileSt.UpLoad(rstOrdMilest)
    ' Комплектование
    Set rstOrdMilest = objOrdMileSt.LoadUPD(0)
    rstOrdMilest!lngOrderID = lngOrderID
    rstOrdMilest!datActDate = colMSDates.Item(3)
    rstOrdMilest!datPlanDate = rstOrdMilest!datActDate
    rstOrdMilest!lngTechMilestID = objOrdMileSt.lngStockMSID
    rstOrdMilest!intSequenceNum = objOrdMileSt.GetTMSequenceByID(rstOrdMilest!lngTechMilestID)
    rstOrdMilest!boolIsFinish = False
    rc = objOrdMileSt.UpLoad(rstOrdMilest)
    ' Доставка
    Set rstOrdMilest = objOrdMileSt.LoadUPD(0)
    rstOrdMilest!lngOrderID = lngOrderID
    rstOrdMilest!datActDate = colMSDates.Item(5)
    rstOrdMilest!datPlanDate = rstOrdMilest!datActDate
    rstOrdMilest!datActFDate = rstOrdMilest!datActDate
    rstOrdMilest!lngTechMilestID = objOrdMileSt.lngShipMSID
    rstOrdMilest!intSequenceNum = objOrdMileSt.GetTMSequenceByID(rstOrdMilest!lngTechMilestID)
    rstOrdMilest!boolIsFinish = False
    rc = objOrdMileSt.UpLoad(rstOrdMilest)
    ' Монтаж
    Set rstOrdMilest = objOrdMileSt.LoadUPD(0)
    rstOrdMilest!lngOrderID = lngOrderID
    rstOrdMilest!datActDate = colMSDates.Item(colMSDates.Count)
    rstOrdMilest!datPlanDate = rstOrdMilest!datActDate
    rstOrdMilest!datActFDate = rstOrdMilest!datActDate
    rstOrdMilest!lngTechMilestID = objOrdMileSt.lngMountMSID
    rstOrdMilest!intSequenceNum = objOrdMileSt.GetTMSequenceByID(rstOrdMilest!lngTechMilestID)
    rstOrdMilest!boolIsFinish = False
    rc = objOrdMileSt.UpLoad(rstOrdMilest)
    
    'состояний Договоров
    Dim objContractDoc As clsContractDoc
    Dim rstContractDoc As Recordset
    Set objContractDoc = New clsContractDoc
    objContractDoc.BCData = objBCData
    Set rstContractDoc = objContractDoc.Load(0)       ' Cоздать новую запись из журнала состояний Договоров
    rstContractDoc!lngOrderID = lngOrderID            ' Связать запись с данным заказом по ID
    rc = objContractDoc.UpLoad(rstContractDoc)        ' Записать новую запись в журнал состояний Договоров
    Set rstContractDoc = Nothing
    Set objContractDoc = Nothing

    CreateOrderByMetReq = lngOrderID
    Set objCurrency = Nothing
    Set objCRate = Nothing
    Set rstCRate = Nothing
    Set objMetReq = Nothing
    Set rstVendor = Nothing
    Set rstMaker = Nothing
    Set objCRate = Nothing
    Set objCodif = Nothing
    Set objUser = Nothing
    Set objProp = Nothing
    Set objOrdMileSt = Nothing
    Set rstOrder = Nothing
    Set rstProps = Nothing
    Set rstProfile = Nothing
    Set rstOrdMilest = Nothing
    Set objDealerData = Nothing
    Set rstDealerData = Nothing
    Set colMSDates = Nothing
    Set objFcdOrder = Nothing
End Function
