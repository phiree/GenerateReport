import pypyodbc
from datetime import datetime
from decimal import Decimal
import xlwt3
import os
import sys
import FitSheetWrapper
import config

class ReportGenerator:
    def __init__(self, date_start, date_end,reportname):
        self.date_start = date_start
        self.date_end = date_end
        self.reportname=reportname
    def generate_reports(self):
        sale_bill_detail='''
        (
				select TStockIOBill.JID,TStockIOBill.JDeptID,TStockIOBill.JSupClientID,

				TStockIOBill.JHandleID,TStockIOBill.JMemo,TStockIOGrid.JGoodsID,TStockIOBill.JBillDate,
				TStockIOBill.JBillCode,TStockIOBill.JBillType,JGoodsQty=TStockIOGrid.JGridQty,
				JGoodsPrice=TStockIOGrid.JGridPrice*TStockIOGrid.JDiscRate,
				JGoodsAmt= case when TStockIOBill.jbilltype=1199 then -1 else 1 end * TStockIOGrid.JGridAmt,JCollectAmt=0.0
				from TStockIOBill
				left outer join TStockIOGrid on TStockIOBill.JID=TStockIOGrid.JBillID

				where TStockIOBill.JUseID>=0 and TStockIOBill.JBillType in (1201,1199) --JuseID: -1表示已取消 0 表示与其他单据没关联 大于零表示关联的ID,
				union all
				--------调价单
				select JID,JDeptID,JSupClientID,JHandleID,JMemo,JGoodsID=0,JBillDate,JBillCode,JBillType,JGoodsQty=0.0,
				JGoodsPrice=0.0,JGoodsAmt=JBillAmt,JCollectAmt=0.0
				from TAdjBill where TAdjBill.JUseID>=0 and TAdjBill.JBillType=1402
				union all
				select TPosBill.JID,JDeptID=TStock.JDeptID,
				JSupClientID,JHandleID,JMemo,JGoodsID,JBillDate,
				JBillCode,JBillType,JGoodsQty,
				JGoodsPrice,JGoodsAmt,JCollectAmt=0.0
				from
					(
					-------- 未转移的 零售单据
					select TPosBill.JID,JBillType=1207,JStockID,TPosBill.JSupClientID,TPosBill.JHandleID,JBillDate,
					JBillCode=CONVERT(varchar(20),JSequenceID),
					JMemo,TPosGrid.JGoodsID,JGoodsQty=TPosGrid.JGridQty,JGoodsPrice=TPosGrid.JPointSalePrice,
					JGoodsAmt=TPosGrid.JGridAmt
					from TPosBill
					left outer join TPosGrid on TPosGrid.JBillID=TPosBill.JID

					union all
					---------- 已转移的 零售单据
					select TPosBillHist.JBillID,JBillType=1207,JStockID,TPosBillHist.JSupClientID,TPosBillHist.JHandleID,JBillDate,
					JBillCode=CONVERT(varchar(20),JSequenceID),
					JMemo,TPosGridHist.JGoodsID,JGoodsQty=TPosGridHist.JGridQty,JGoodsPrice=TPosGridHist.JPointSalePrice,
					JGoodsAmt=TPosGridHist.JGridAmt
					from TPosBillHist
					left outer join TPosGridHist on TPosGridHist.JBillID=TPosBillHist.JBillID
					) as TPosBill
				left outer join TStock on TStock.JID=TPosBill.JStockID
				) as a
        '''
        sale_product_detail='''
        (
            select  a.JGoodsID, b.jgoodscode as Product_Code,b.jclassid,
            b.jgoodsname as Product_Name,
            sum(a.jgoodsamt) as Total_Amount,
            b.jrefcostprice as Reference_Cost_Price,
            b.jrefcostprice*sum(a.jgoodsqty) as Reference_Cost_Amount,
            sum(a.jgoodsamt)- b.jrefcostprice*sum(a.jgoodsqty) as Reference_Profit,
            case sum(a.jgoodsamt) when 0 then 0 else (sum(a.jgoodsamt)- b.jrefcostprice*sum(a.jgoodsqty))/sum(a.jgoodsamt) end as Reference_Profit_Rate,
            sum(a.jgoodsqty) as Total_Sale_Quantity
            from 
            '''+sale_bill_detail+'''
            inner join tgoods b
            on a.jgoodsid=b.jid
            where a.jbilldate between '{0}' and dateadd(d,1,'{1}')
            group by b.jgoodsname,b.jgoodscode,b.jrefcostprice,a.JGoodsID,b.JClassID
          ) 
         as product_sale_summary
        '''.format(self.date_start, self.date_end)
        list_sql = (
            ('Bill_Report',
            [
                ('Bill_Relation_Purchase',
                 '''
                 select tt1.englishname as Bill_Name,a.jbillcode as Bill_Code,
                 'Qty:',a.jgridqty as Quantity,a.jbillamt as Bill_Amount,a.jbilldate as Bill_Date,
                 tt2.englishname as Bill_Name,b.jbillcode as Bill_Code,b.jbilldate as Bill_Date,'Qty:',b.jgridqty as Quantity
                 ,b.jbillamt as Bill_Amount,
                 tt3.englishname as Bill_Name,b.jrelationfrom as Bill_Code,f.jbilldate as Bill_Date,'Qty:'
                 ,sum(og.jgridqty) as Quantity,f.jbillamt as Bill_Amount from VStockIOBillBrow a
                 left join VStockIOBillBrow b
                 on a.juseid=b.jid and a.jusebilltype=b.jbilltype
                 left join torderbill f on f.jid=b.juseid
                 left join tbillinfo c on a.jbilltype=c.jid left join temp_english_name_billtype tt1 on tt1.typeid=c.jid
                 left join tbillinfo d on b.jbilltype=d.jid left join temp_english_name_billtype tt2 on tt2.typeid=d.jid
                 left join tbillinfo e on b.jusebilltype=e.jid left join temp_english_name_billtype tt3 on tt3.typeid=e.jid
                 left join tordergrid og on og.jbillid=f.jid
                 where b.jrelationfrom is not null
                    and a.jbilldate between '{0}' and dateadd(d,1,'{1}')
                    and a.jbilltype=1101
                 group by f.jid, tt1.englishname,a.jbillcode,
                 --'Qty',
                 a.jgridqty,a.jbillamt,a.jbilldate,
                 tt2.englishname,b.jbillcode,b.jbilldate,
                 --'Qty',
                 b.jgridqty,b.jbillamt,
                 tt3.englishname,b.jrelationfrom,f.jbilldate,f.jbillamt
                 order by a.jbilldate desc
                 --,'Qty'
                 '''.format(self.date_start, self.date_end)
                ),
                ('Bill_Relation_Sale',
                 '''
                 select tt1.englishname as Bill_Name,a.jbillcode as Bill_Code,'Qty:',a.jgridqty as Quantity,a.jbillamt as Bill_Amount
                 ,a.jbilldate as Bill_Date,
                 tt2.englishname as Bill_Name,b.jbillcode as Bill_Code,b.jbilldate as Bill_Date,'Qty:',b.jgridqty as Quantity
                 ,b.jbillamt as Bill_Amount,
                 tt3.englishname as Bill_Name,b.jrelationfrom as Bill_Code,f.jbilldate as Bill_Date,'Qty:',sum(og.jgridqty) as Quantity
                 ,'Amt:',f.jbillamt as Bill_Amount from VStockIOBillBrow a
                 left join VStockIOBillBrow b
                 on a.juseid=b.jid and a.jusebilltype=b.jbilltype
                 left join torderbill f on f.jid=b.juseid
                 left join tbillinfo c on a.jbilltype=c.jid left join temp_english_name_billtype tt1 on tt1.typeid=c.jid
                 left join tbillinfo d on b.jbilltype=d.jid left join temp_english_name_billtype tt2 on tt2.typeid=d.jid
                 left join tbillinfo e on b.jusebilltype=e.jid left join temp_english_name_billtype tt3 on tt3.typeid=e.jid
                 left join tordergrid og on og.jbillid=f.jid
                 where b.jrelationfrom is not null
                    and b.jbilldate between '{0}' and dateadd(d,1,'{1}')
                    and b.jbilltype=1201
                 group by f.jid, tt1.englishname,a.jbillcode,
                 --'Qty',
                 a.jgridqty,a.jbillamt,a.jbilldate,
                 tt2.englishname,b.jbillcode,b.jbilldate,
                 --'Qty',
                 b.jgridqty,b.jbillamt,
                 tt3.englishname,b.jrelationfrom,f.jbilldate,f.jbillamt
                 order by a.jbilldate desc
                 --,'Qty'
                 '''.format(self.date_start, self.date_end)
                )

            ])
           
            ,
            ('Sale_Report',
            [
                ('Product_Sale',
                
                 '''
                  select product_sale_summary.Product_Code, 
		product_sale_summary.Product_Name,
		product_sale_summary.Total_Amount,
		product_sale_summary.Reference_Cost_Price,
		product_sale_summary.Reference_Cost_Amount,
		product_sale_summary.Reference_Profit,
		product_sale_summary.Reference_Profit_Rate,
		importamount.total_import_cp as Total_WareEntry_Quantity_CP,-- total warehouseentry amount bwtween selected period.
		product_sale_summary.Total_Sale_Quantity,
    importamount.totalimport as Total_Import_Quantity,Total_Sale_Quantity/case totalimport when 0 then null else totalimport end  as Move_Rate
    from '''+sale_product_detail+'''
		left join 
 (
			select jgoodsid, SUM(total_qty) as  totalimport from
			(
			--入库-出库
			select  b.jgoodsid,  sum(jgridqty)* case when a.JBillType in (1102,1104) then 1 else -1 end  as total_qty  
			from tstockiobill  a  
			inner join tstockiogrid b
			on a.jid=b.jbillid  
			where 1=1
			and a.jbilltype  in  (1102,1104,1204)
			group by b.jgoodsid,a.JBillType
			) as t_sum_qty
			left join (
			select  b.jgoodsid, case when a.jbilltype in (1102,1104) then sum(jgridqty) end as total_import ,
			 a.jbilltype 
			from tstockiobill  a  
			inner join tstockiogrid b
			on a.jid=b.jbillid  
			where 1=1 
			and a.jbilldate  '{0}' and dateadd(d,1,'{1}')
			and a.jbilltype  in  (1102)
			group by b.jgoodsid,a.JBillType
			) as t_sum_qty_cp
			on t_sum_qty.jgoodsid=t_sum_qty_cp.jgoodsid and t_sum_qty.jbilltype=t_sum_qty_cp.jbilltype

			group by (t_sum_qty.JGoodsID)	
	)
as importamount
on product_sale_summary.jgoodsid=importamount.jgoodsid 
               
                order by Total_Amount desc '''.format(self.date_start, self.date_end)
                )
                ,
                ('Product_Category_Sale',
                 '''
select category_sale_summary.JClassCode as Category_Code
,category_sale_summary.JClassName as Category_Name
,category_sale_summary.Total_Sale_Amount
,category_sale_summary.Total_Cost_Amount
,category_sale_summary.Total_Sale_Quantity
,import_total.totalimport as Total_Import_Quantity
,category_sale_summary.Total_Sale_Quantity/case totalimport when 0 then null else totalimport end  as Move_Rate
from 
(
select JClassCode,JClassName,SUM(product_sale_summary.Total_Amount) as Total_Sale_Amount
,SUM(product_sale_summary.Reference_Cost_Amount) as Total_Cost_Amount,
SUM(product_sale_summary.Total_Sale_Quantity) as Total_Sale_Quantity
from
  '''+sale_product_detail+'''
  left join tbasicsort s on product_sale_summary.jclassid=s.jid
	group by s.JClassName,s.JClassCode
)
as category_sale_summary
left join (
		select t_sum_qty.JClassCode,t_sum_qty.JClassName, SUM(total_qty) as  totalimport from
			(
			--入库
			select  s.JClassCode,s.JClassName,  sum(jgridqty)* case when a.JBillType in (1102,1104) then 1 else -1 end  as total_qty
			from tstockiobill  a
			inner join tstockiogrid b
			on a.jid=b.jbillid
			inner join TGoods g on b.JGoodsID=g.JID
			left join TBasicSort s on s.JID=g.JClassID
			where  a.jbilltype  in  (1102,1104,1204)
			group by s.JClassCode,s.JClassName,a.JBillType
			
			) as t_sum_qty
		group by JClassName,JClassCode
        ) as import_total
on import_total.JClassCode=category_sale_summary.JClassCode
                    '''
                )
                ,(
                'Product_Never_Sold',
                '''
                select jgoodscode as Code,jgoodsname as  Name,jmodel as Factory_Code,jpointsaleprice as Sale_Price,jrefcostprice as Cost_Price
                from tgoods where jgoodscode not in(select jgoodscode from VSaleAccCompare) order by jgoodscode
                '''
                )
                ])
            ,('Customer_Report',
                [
                ('Customer_Product',
                 '''
                 -- details for customer
                 select case when  b.jsupclientname is null then 'General(Cash)' else b.jsupclientname end as Customer_Name,g.jgoodscode as Product_Code,g.jgoodsname as Product_Name,sum(a.jgoodsqty) as Total_Quantity,sum(a.jgoodsamt) as Total_Amount
                 from'''+sale_bill_detail+'''
                 inner join tgoods g on a.jgoodsid=g.jid
                 left join tsupclient b on a.jsupclientid=b.jid
                 
                 group by b.jsupclientname,g.jgoodsname,g.jgoodscode
                 order by jsupclientname ,Total_Amount desc
                 '''.format(self.date_start, self.date_end)
                )
                ,
                ('Customer_Product_Category',
                 '''
                 --  summary by sort
                 select case when  b.jsupclientname is null then 'General(Cash)' else b.jsupclientname end as Customer_Name,d.jclassname as Category_Name,sum(a.jgoodsqty) as Total_Quantity,sum(a.jgoodsamt) as Total_Amount
                 from '''+sale_bill_detail+'''
                 left join tsupclient b on a.jsupclientid=b.jid
                 inner join tgoods c on a.jgoodsid=c.jid
                 left join tbasicsort d on c.jclassid=d.jid
                 where jbilldate between '{0}' and dateadd(d,1,'{1}') --and jsupclientname
                 group by b.jsupclientname,d.jclassname
                 order by Total_Amount desc'''.format(self.date_start, self.date_end)
                )
                ,
                ('Customer_Summary',
                 '''
                 --  summary by customer
                 select case when  b.jsupclientname is null then 'General(Cash)' else b.jsupclientname end as Customer_Name,sum(a.jgoodsqty) as Total_Quantity,sum(a.jgoodsamt) as Total_Amount
                 from '''+sale_bill_detail+'''
                 left join tsupclient b on a.jsupclientid=b.jid
                 where jbilldate between '{0}' and dateadd(d,1,'{1}')  --and jsupclientname
                 group by b.jsupclientname
                 order by Total_Amount desc'''.format(self.date_start, self.date_end)
                ),
                ('Customer_Info',
                 '''
                select jsupclientcode as Code,jsupclientName as Name,jaddress as Address,jPostcode as Postcode,
                jcontact as Contact,jTele as Tel,jmobilenumber as Mobile,jfax as Fax,jEmail as Email,jwebsite as Website,
                jcountry as Country,jcompany as Company,jcity as City,jStartdaten as Create_Date
                from tsupclient
                where jfunctionid=30700 and jnouse=0--means customer
                order by name'''
                ),
            ]
            ),
            ('Account_Receivable_Report',
            [
                ('Receivable_Summary',
                 '''
                select   d.jname as Bill_Name,case when a.iscanceled<0 then 'Canceled' else '' end as Is_Bill_Canceled
                ,   f.jsupclientname as Customer_Name,a.jbillcode as Bill_Code
                , a.jbilldate as Bill_Date,a.jbillamt as Receivable_Amount,a.OriginalBillAmt as Original_Bill_Amount,
                case when sum(b.jcurclramt) is null then 0 else sum(b.jcurclramt) end as Paid_Amount
                ,a.jbillamt-case when sum(b.jcurclramt) is null then 0 else sum(b.jcurclramt) end as  Not_Paid_Amount
                from
                (
                /*挂账零售单据 只当作应付款单据,不需要当作收款单据.
                排除取消的销售单据(只有销售开票有取消状态)
                */
                select   a.jbillid ,DATEADD(dd, 0, DATEDIFF(dd, 0, b.jbilldate)) as jbilldate
                    ,b.jbilldate as fulljbilldate,jsupclientid,1207 as jbilltype
                    ,a.jcollectamt as jbillamt,b.jbillamt as OriginalBillAmt
                    ,jcustombillcode as jbillcode,0 as iscanceled
                from TPosCollectGridHist a
                inner join  TPosBillHist b on a.jbillid=b.jbillid
                where a.jcollectamt>0
                union all
                select jid as jbillid,DATEADD(dd, 0, DATEDIFF(dd, 0,jbilldate)) as jbilldate
                    ,jbilldate as fulljbilldate,jsupclientid, jbilltype,jbillamt
                    , jbillamt as OriginalBillAmt, jbillcode,case when juseid>=0 then 0 else -1 end  as iscanceled
                from tstockiobill where jbilltype in (1201,1207)
                ) a           ---应收账款
                left join TPayCollectGrid b --付款单

                on a.jbillid=b.juseid and b.jusebilltype=a.jbilltype
                left join
                TPayCollectBill
                c  on b.jbillid=c.jid and c.juseid>=0  -- 付款单总表
                inner join tbillinfo d on d.jid=a.jbilltype
                left join tbillinfo e on e.jid=c.jbilltype
                inner join tsupclient f on f.jid=a.jsupclientid
                where a.jbilltype in (1201,1207) and a.jbilldate  between '{0}' and dateadd(d,1,'{1}')
                    and c.juseid>=0
                group by d.jname ,f.jsupclientname,a.jbillcode , a.jbilldate ,a.jbillamt ,a.OriginalBillAmt,a.fulljbilldate,
c.juseid,a.iscanceled
                order by a.iscanceled desc ,a.fulljbilldate desc
                 '''.format(self.date_start, self.date_end)
                ),
                ('Receivable_Detail',
                 '''
                select d.jname as Bill_Name,case when a.juseid<0 then 'Canceled'  else '' end as Is_Bill_Canceled
                , f.jsupclientname as Customer_Name,a.jbillcode as Bill_Code, a.jbilldate as Bill_Date
                ,a.jbillamt as Receivable_Amount,a.OriginalBillAmt as Original_Bill_Amount ,
                e.jname as Bill_Name,case when c.juseid<0 then 'Canceled'  else '' end as Is_Pay_Canceled,
				 c.jbillcode as Bill_Code ,c.jbilldate as Bill_Date,b.jcurclramt as Paid_Amount
                from
                (
                select a.jbillid ,DATEADD(dd, 0, DATEDIFF(dd, 0, b.jbilldate)) as jbilldate,b.jbilldate as fulljbilldate, jsupclientid,1207 as jbilltype,a.jcollectamt as jbillamt,b.jbillamt as OriginalBillAmt,jcustombillcode as jbillcode,0 as juseid  from TPosCollectGridHist a inner join  TPosBillHist b on a.jbillid=b.jbillid
                where a.jcollectamt>0
                union all
                select jid as jbillid,DATEADD(dd, 0, DATEDIFF(dd, 0,jbilldate)) as jbilldate,jbilldate as fulljbilldate,jsupclientid, jbilltype,jbillamt, jbillamt as OriginalBillAmt, jbillcode,juseid   from tstockiobill where jbilltype in (1201,1207)
                ) a                                                 ---应收账款
                left join TPayCollectGrid b --付款单

                on a.jbillid=b.juseid and b.jusebilltype=a.jbilltype
                left join
                TPayCollectBill
                c  on b.jbillid=c.jid  -- 付款单总表
                inner join tbillinfo d on d.jid=a.jbilltype
                left join tbillinfo e on e.jid=c.jbilltype
                inner join tsupclient f on f.jid=a.jsupclientid
                where a.jbilltype in (1201,1207)
                and a.jbilldate  between '{0}' and dateadd(d,1,'{1}')
                order by a.juseid desc, a.fulljbilldate desc,a.jbillcode desc
                 '''.format(self.date_start, self.date_end)
                )

            ])
        )
        for item in list_sql:
            if self.reportname!=item[0]:
                continue
            self.create_excel_book(item[0], item[1])
        pass

    def create_excel_book(self, bookname, list_sql):
        wb = xlwt3.Workbook(style_compression=2)
        for counter, sql in enumerate(list_sql):
            self.add_sheet_excel(sql[1], wb, sql[0])
        report_folder=os.getcwd()+'\\reports\\'
        if not os.path.exists(report_folder):
            os.makedirs(report_folder)
        wb.save(report_folder+bookname + '_{0}__{1}.xls'.format(self.date_start, self.date_end))

    def add_sheet_excel(self, sql, wb, sheetname):
        
        report_table = self.getdata(sql)
        
        ws = wb.add_sheet(sheetname)
        ws.set_panes_frozen(True) # frozen headings instead of split panes
        ws.set_horz_split_pos(1) # in general, freeze after last heading
        wrap_ws = FitSheetWrapper.FitSheetWrapper(ws)
        self.createSheet(report_table, wrap_ws)

    def createSheet(self, table_data, ws):
        mystyle = xlwt3.easyxf('''borders: left thin, right thin, top thin, bottom thin;
    pattern: pattern solid, fore_colour yellow;''')
        for counter_row, row_data in enumerate(table_data):
            for counter_column, field_data in enumerate(row_data):
                if counter_row == 0:
                    ws.write(counter_row, counter_column, field_data.title(), mystyle)
                else:
                    cell_style = xlwt3.XFStyle()
                    if isinstance(field_data,datetime):
                        cell_style.num_format_str = 'dd/mm/yyyy'
                    ws.write(counter_row, counter_column, field_data,cell_style)


    def getdata(self, sql):

        print(sql)

        conn = pypyodbc.connect(config.connection_string)
        
        cursor = conn.cursor()

        cursor.execute(sql)
        report_table = cursor.fetchall()

        columns = [c[0] for c in cursor.description]
        report_table.insert(0, columns)
        conn.commit()
        conn.close()
        return report_table


if __name__ == "__main__":
    date_start = sys.argv[1]
    date_end = sys.argv[2]
    ReportGenerator(date_start, date_end).generate_reports()