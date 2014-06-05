import pypyodbc
from datetime import datetime
from decimal import Decimal
import xlwt3
import sys
import FitSheetWrapper


class ReportGenerator:
    def __init__(self, date_start, date_end,reportname):
        self.date_start = date_start
        self.date_end = date_end
        self.reportname=reportname
    def generate_reports(self):
        list_sql = (
            ('BillReport',
            [
                ('Bill_Relation_Purchase',
                 '''
                 select tt1.englishname,a.jbillcode,'Qty',a.jgridqty,a.jbillamt,a.jbilldate,
                 tt2.englishname,b.jbillcode,b.jbilldate,'Qty',b.jgridqty,b.jbillamt,
                 tt3.englishname,b.jrelationfrom,f.jbilldate,'Qty',sum(og.jgridqty),'Amt',f.jbillamt from VStockIOBillBrow a
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
                 select tt1.englishname,a.jbillcode,'Qty',a.jgridqty,a.jbillamt,a.jbilldate,
                 tt2.englishname,b.jbillcode,b.jbilldate,'Qty',b.jgridqty,b.jbillamt,
                 tt3.englishname,b.jrelationfrom,f.jbilldate,'Qty',sum(og.jgridqty),'Amt',f.jbillamt from VStockIOBillBrow a
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
            ('SaleReport',
            [
                ('Product Sale Report',
                 '''
                select a.jgoodscode as Product_Code,
                a.jgoodsname as Product_Name,
                sum(a.jgoodsqty) as Total_Quantity,
                sum(a.jgoodsamt) as Total_Amount,
                b.jrefcostprice as Reference_Cost_Price,
                b.jrefcostprice*sum(a.jgoodsqty) as Reference_Cost_Amount,
                sum(a.jgoodsamt)- b.jrefcostprice*sum(a.jgoodsqty) as Reference_Profit,
                case sum(a.jgoodsamt) when 0 then 0 else (sum(a.jgoodsamt)- b.jrefcostprice*sum(a.jgoodsqty))/sum(a.jgoodsamt) end as Reference_Profit_Rate
                from VSaleAccCompare  a
                inner join tgoods b 
                on a.jgoodscode=b.jgoodscode
                where 1=1 
                and jbilldate between '{0}' and dateadd(d,1,'{1}') --and jsupclientname
                group by a.jgoodsname,a.jgoodscode,b.jrefcostprice
                order by Total_Amount desc
                 '''.format(self.date_start, self.date_end)
                )
                ,
                ('Product Category Sale Report',
                 '''
                    -- product sort sale report
                    select d.jclasscode as Category_Code,d.jclassname as Category_Name,sum(a.jgoodsqty) as Total_Quantity,sum(a.jgoodsamt) as Total_Amount
                    from VSaleAccCompare  a
                    inner join tgoods c on a.jgoodsid=c.jid
                    inner join tbasicsort d on c.jclassid=d.jid
                    where jbilldate between '{0}' and dateadd(d,1,'{1}') --and jsupclientname
                    group by d.jclasscode,d.jclassname
                    order by Total_Amount desc
                    '''.format(self.date_start, self.date_end)
                )
                ,(
                'Product Never Sold',
                '''
                select jgoodscode as Code,jgoodsname as  Name,jmodel as Factory_Code,jpointsaleprice as Sale_Price,jrefcostprice as Cost_Price
                from tgoods where jgoodscode not in(select jgoodscode from VSaleAccCompare) order by jgoodscode
                '''
                )
                ])
            ,('CustomerReport',
                [
                ('Customer_Product Report',
                 '''
                 -- details for customer
                 select b.jsupclientname as Customer_Name,a.jgoodscode as Product_Code,a.jgoodsname as Product_Name,sum(a.jgoodsqty) as Total_Quantity,sum(a.jgoodsamt) as Total_Amount
                 from VSaleAccCompare  a
                 inner join tsupclient b on a.jsupclientid=b.jid
                 where jbilldate between '{0}' and dateadd(d,1,'{1}') --and jsupclientname
                 group by b.jsupclientname,a.jgoodsname,a.jgoodscode
                 order by jsupclientname ,Total_Amount desc
                 '''.format(self.date_start, self.date_end)
                )
                ,
                ('Customer_ProductCategory Report',
                 '''
                 --  summary by sort
                 select b.jsupclientname as Customer_Name,d.jclassname as Category_Name,sum(a.jgoodsqty) as Total_Quantity,sum(a.jgoodsamt) as Total_Amount
                 from VSaleAccCompare  a
                 inner join tsupclient b on a.jsupclientid=b.jid
                 inner join tgoods c on a.jgoodsid=c.jid
                 inner join tbasicsort d on c.jclassid=d.jid
                 where jbilldate between '{0}' and dateadd(d,1,'{1}') --and jsupclientname
                 group by b.jsupclientname,d.jclassname
                 order by Total_Amount desc'''.format(self.date_start, self.date_end)
                )
                ,
                ('Customer Summary Report',
                 '''
                 --  summary by customer
                 select b.jsupclientname as Customer_Name,sum(a.jgoodsqty) as Total_Quantity,sum(a.jgoodsamt) as Total_Amount
                 from VSaleAccCompare  a
                 inner join tsupclient b on a.jsupclientid=b.jid
                 where jbilldate between '{0}' and dateadd(d,1,'{1}')  --and jsupclientname
                 group by b.jsupclientname
                 order by Total_Amount desc'''.format(self.date_start, self.date_end)
                ),
                ('Customer Info',
                 '''
                select jsupclientcode as Customer_Code,jsupclientName as Name,jaddress as Address,jPostcode as Postcode,
                jcontact as ContactPerson,jTele as Tel,jmobilenumber as Mobile,jfax as Fax,jEmail as Email,jwebsite as Website,
                jcountry as Country,jcompany as CompanyName,jcity as City,jStartdaten as Create_Date
                from tsupclient
                where jfunctionid=30700 and jnouse=0--means customer
                order by name'''
                ),
            ]
            )
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
        wb.save(bookname + '_{0}__{1}.xls'.format(self.date_start, self.date_end))

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



        conn = pypyodbc.connect(
            "driver={SQL Server};server=server-pc\sqlexpress;database=TSNET1001;uid=ntsmyanmar;pwd=12345678")
        # conn=pypyodbc.connect("driver={SQL Server};server=.\sqlexpress;database=TSNET1013;trusted_connection=yes;")


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