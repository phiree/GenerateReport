select d.jclasscode as Category_Code,d.jclassname as Category_Name,sum(a.jgoodsqty) as TotalQuantity,sum(a.jgoodsamt) as TotalAmount
                    from VSaleAccCompare  a
                    inner join tgoods c on a.jgoodsid=c.jid
                    inner join tbasicsort d on c.jclassid=d.jid
					inner join 
                    group by d.jclasscode,d.jclassname
                    order by totalAmount desc

select * from VSaleAccCompare  a where  cast(a.jgoodsqty as int) <> a.jgoodsqty

sp_helptext vsaleacccompare

select d.jbillcode,* from tstockiogrid a 
inner join tgoods b 
on a.jgoodsid=b.jid
inner join tbasicsort c
on c.jid=b.jclassid
inner join tstockiobill d
on d.jid=a.jbillid
where --c.jclasscode='099'
cast(a.jgridqty as int) <> a.jgridqty




select * from tstockiobill where jid=219


Text
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


--应收账款的本期减少项增加POS挂账的现收功能
CREATE View VSaleAccCompare    
 as select TSaleAccCompare.JID,TSaleAccCompare.JBillType,TSaleAccCompare.JDeptID,    
 TSaleAccCompare.JSupClientID,TSaleAccCompare.JBillDate,TSaleAccCompare.JBillCode,    
 JMemo=TSaleAccCompare.JMemo,    
 JSummary='['+TBillInfo.JName+']'+case TSaleAccCompare.JHandleID when 0 then ''     
 else '['+TEmployee.JEmpName+']' end+TSaleAccCompare.JMemo,    
 TSaleAccCompare.JGoodsID,JGoodsCode=case TSaleAccCompare.JGoodsID when 0 then '' else TGoods.JGoodsCode end,    
 JGoodsName=case TSaleAccCompare.JGoodsID when 0 then '' else TGoods.JGoodsName end,    
 JModel=case TSaleAccCompare.JGoodsID when 0 then '' else TGoods.JModel end,    
 JUnit=case TSaleAccCompare.JGoodsID when 0 then '' else TGoods.JUnit end,    
 TSaleAccCompare.JGoodsQty,TSaleAccCompare.JGoodsPrice,    
 TSaleAccCompare.JGoodsAmt,TSaleAccCompare.JCollectAmt,JBalanceAmt=0.0    
 from    
     (select TStockIOBill.JID,TStockIOBill.JDeptID,TStockIOBill.JSupClientID,    
     TStockIOBill.JHandleID,TStockIOBill.JMemo,TStockIOGrid.JGoodsID,TStockIOBill.JBillDate,    
     TStockIOBill.JBillCode,TStockIOBill.JBillType,JGoodsQty=TStockIOGrid.JGridQty,    
     JGoodsPrice=TStockIOGrid.JGridPrice*TStockIOGrid.JDiscRate,JGoodsAmt=TStockIOGrid.JGridAmt,JCollectAmt=0.0    
     from TStockIOBill    
     left outer join TStockIOGrid on TStockIOBill.JID=TStockIOGrid.JBillID    
     where TStockIOBill.JUseID>=0 and TStockIOBill.JBillType in (1201,1199)    
   union all    
     select JID,JDeptID,JSupClientID,JHandleID,JMemo,JGoodsID=0,JBillDate,JBillCode,JBillType,JGoodsQty=0.0,    
     JGoodsPrice=0.0,JGoodsAmt=JBillAmt,JCollectAmt=0.0    
     from TAdjBill where TAdjBill.JUseID>=0 and TAdjBill.JBillType=1402    
   union all    
		select TPosBill.JID,JDeptID=TStock.JDeptID,                 
		JSupClientID,JHandleID,JMemo,JGoodsID,JBillDate,                 
		JBillCode,JBillType,JGoodsQty,                 
		JGoodsPrice,JGoodsAmt,JCollectAmt=0.0                 
	    from                  
	       (select TPosBill.JID,JBillType=1207,JStockID,TPosBill.JSupClientID,TPosBill.JHandleID,JBillDate,                 
 			JBillCode=CONVERT(varchar(20),JSequenceID),                   
				JMemo,TPosGrid.JGoodsID,JGoodsQty=TPosGrid.JGridQty,JGoodsPrice=TPosGrid.JPointSalePrice,                 
   			JGoodsAmt=TPosGrid.JGridAmt                    
   			 from TPosBill                    
   			left outer join TPosGrid on TPosGrid.JBillID=TPosBill.JID                 
   			where TPosBill.JSupClientID >0                 
   			union all                    
   			select TPosBillHist.JBillID,JBillType=1207,JStockID,TPosBillHist.JSupClientID,TPosBillHist.JHandleID,JBillDate,                 
   			JBillCode=CONVERT(varchar(20),JSequenceID),                   
   			JMemo,TPosGridHist.JGoodsID,JGoodsQty=TPosGridHist.JGridQty,JGoodsPrice=TPosGridHist.JPointSalePrice,                 
   			JGoodsAmt=TPosGridHist.JGridAmt                    
   			 from TPosBillHist                    
   			left outer join TPosGridHist on TPosGridHist.JBillID=TPosBillHist.JBillID                 
   			where TPosBillHist.JSupClientID >0                 
   			) as TPosBill                   
	     left outer join TStock on TStock.JID=TPosBill.JStockID  
   union all    
      select b.JID,JDeptID=TStock.JDeptID,                 
		JSupClientID,JHandleID,JMemo,JGoodsID=0,JBillDate,                 
		JBillCode=CONVERT(varchar(20),JSequenceID),JBillType=1207,JGoodsQty=0,                 
		JGoodsPrice=0,JGoodsAmt=0,JCollectAmt=b.JBillAmt-c.JSumCollectAmt                      
		 from TPosBill b
		  LEFT OUTER JOIN (SELECT JBillID,JSumCollectAmt=SUM(JCollectAmt) 
	                 FROM TPosCollectGrid 
	                 WHERE JSettlementID =(SELECT JID FROM TPosSettlement WHERE JMarkID=2)
	                 GROUP BY JBillID) c
	      ON b.JID=c.JBillID 
	      left outer join TStock on TStock.JID=b.JStockID                   
		where b.JSupClientID >0            
	union all                    
	    select b.JID,JDeptID=TStock.JDeptID,                 
		JSupClientID,JHandleID,JMemo,JGoodsID=0,JBillDate,                 
		JBillCode=CONVERT(varchar(20),JSequenceID),JBillType=1207,JGoodsQty=0,                 
		JGoodsPrice=0,JGoodsAmt=0,JCollectAmt=b.JBillAmt-c.JSumCollectAmt               
		 from TPosBillHist b
		  LEFT OUTER JOIN (SELECT JBillID,JSumCollectAmt=SUM(JCollectAmt) 
	                 FROM TPosCollectGridHist 
	                 WHERE JSettlementID =(SELECT JID FROM TPosSettlement WHERE JMarkID=2)
	                 GROUP BY JBillID) c
	      ON b.JBillID=c.JBillID     
	       left outer join TStock on TStock.JID=b.JStockID               
		where b.JSupClientID >0         
		 
   union all    
     select JID,JDeptID,JSupClientID,JHandleID,JMemo,JGoodsID=0,JBillDate,JBillCode,JBillType,JGoodsQty=0.0,    
     JGoodsPrice=0.0,JGoodsAmt=0.0,JCollectAmt=JBillAmt    
     from TPayCollectBill where JUseID>=0 and JBillType=1502    
   union all    
     select TPayCollectBill.JID,TPayCollectBill.JDeptID,TStockIOBill.JSupClientID,    
     TPayCollectBill.JHandleID,TPayCollectBill.JMemo,JGoodsID=0,TPayCollectBill.JBillDate,    
     TPayCollectBill.JBillCode,TPayCollectBill.JBillType,JGoodsQty=0.0,JGoodsPrice=0.0,JGoodsAmt=0.0,    
     JCollectAmt=sum(TPayCollectGrid.JCurClrAmt+TPayCollectGrid.JCurDiscAmt+TPayCollectGrid.JCurExpensesAmt)    
     from TPayCollectBill     
     left outer join TPayCollectGrid on TPayCollectGrid.JBillID=TPayCollectBill.JID    
     left outer join TStockIOBill on TStockIOBill.JID=TPayCollectGrid.JUseID     
     and TStockIOBill.JBillType=TPayCollectGrid.JUseBillType and TStockIOBill.JTransportID=TPayCollectBill.JSupClientID    
     where TPayCollectGrid.JUseID>0 and TPayCollectBill.JBillType=1502 and TStockIOBill.JTransportID>0    
     group by TPayCollectBill.JID,TPayCollectBill.JDeptID,TStockIOBill.JSupClientID,    
     TPayCollectBill.JHandleID,TPayCollectBill.JMemo,TPayCollectBill.JBillDate,    
     TPayCollectBill.JBillCode,TPayCollectBill.JBillType)    
 as TSaleAccCompare    
     left outer join TGoods on TSaleAccCompare.JGoodsID=TGoods.JID                
     left outer join TBillInfo on TSaleAccCompare.JBillType=TBillInfo.JID    
     left outer join TEmployee on TEmployee.JID=TSaleAccCompare.JHandleID


