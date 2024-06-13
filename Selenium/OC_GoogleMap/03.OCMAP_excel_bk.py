# -*- coding: utf-8 -*-
"""
Created on Fri Feb  3 10:42:37 2023

@author: admin
"""


#%%
def ALEXY_All(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
        
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #temp2 where OC = 'ALEXY'
        union all
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'ALEXY'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def ALEXY_TSB(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
        
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'ALEXY'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def ANDERSON_All(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
        
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #temp2 where OC = 'ANDERSON'
        union all
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'ANDERSON'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def ANDERSON_TSB(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
        
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'ANDERSON'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src


def BEN4423_All(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
               
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #temp2 where OC = 'BEN4423'
        union all
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'BEN4423'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def BEN4423_TSB(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
               
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'BEN4423'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src



def BOO5056_All(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
        
       
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #temp2 where OC = 'BOO5056'
        union all
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'BOO5056'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def BOO5056_TSB(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
        
        
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'BOO5056'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def JACKYH_All(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
        
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #temp2 where OC = 'JACKYH'
        union all
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'JACKYH'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def JACKYH_TSB(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
    
        
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'JACKYH'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def JASON4703_All(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
        
        
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #temp2 where OC = 'JASON4703'
        union all
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'JASON4703'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def JASON4703_TSB(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
           
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where oc = 'JASON4703'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def JULIAN_All(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
        
        
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #temp2 where OC = 'JULIAN'
        union all
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'JULIAN'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src



def JULIAN_TSB(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''

        
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'JULIAN'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def SCOTT4162_All(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
        
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #temp2 where OC = 'SCOTT4162'
        union all
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'SCOTT4162'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def SCOTT4162_TSB(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
        
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'SCOTT4162'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def SIMONSH_All(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
    
        
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #temp2 where OC = 'SIMONSH'
        union all
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'SIMONSH'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def SIMONSH_TSB(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
    
        
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'SIMONSH'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def SPANELY_All(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
        
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #temp2 where OC = 'SPANELY'
        union all
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'SPANELY'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def SPANELY_TSB(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
        
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'SPANELY'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src


def WHITE5082_All(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
        
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #temp2 where OC = 'WHITE5082'
        union all
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'WHITE5082'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src

def WHITE5082_TSB(server,username,password,database):
    import pymssql
    conn = pymssql.connect(server=server, user=username, password=password, database = database)
    cursor = conn.cursor()
    script = f"""
    --DELLIST
        select 
        casei = [case],
        v_date,
        oc,
        case_sv_no
        into #DELLIST
        from [10.90.0.194].[OCAP].dbo.outbound_report 
        where CONVERT(varchar(100), v_date, 23)>= CONVERT(varchar(100), getdate(), 23)
        
    --SingleVisit
        select 
        CONVERT(varchar(100), GETDATE(), 23) as Data_date,
        SV_NO as no,
        SV_NO as CaseI,
        SV_Person_Name as CM_Name,
        '台新國際商業銀行股份有限公司' as Bank,
        Amount_TTL as Debt_AMT,
        Case when Case_Type like '%急件%' then '4' else '5' end as PrioritySV,
        'New' as SV_Type,
        '台新單項委外案件' as Case_Type,
        CONVERT(varchar(100), SV_Day, 23) as prioritydate,
        SUBSTRING(ZipCode_GIS,1,3) as ZIP,
        City_GIS as City,
        Town_GIS as Town,
        Address,
        Longitude,
        Latitude,
        SV_Time_Period as Memo,
        Case when Case_Type is null then '' else Case_Type end as Priority,
        '' as Objective,
        UCS_AC_YN as Motivation,
        OC_CN_FST_EMPI as ocempi,
        OC_Code as OC,
        Undertaker_Name as AA,
        Contact_Number as AA_contact,
        case when Status is null and SV_Schedule_Date is not null then CONVERT(varchar(100), SV_Schedule_Date, 23) else '' end as cesv_TIME,
        case when Status is null and SV_Schedule_Date is not null then DATEDIFF(D,CONVERT(varchar(100), GETDATE(), 23),CONVERT(varchar(100), SV_Schedule_Date, 23)) else '' end as countdown,
        CASE WHEN Status is null then '' 
             when Status like '%RECALL%' then 'Recall'
        	 when Status like '%Recall%' then 'Recall'
        	 when Status like '%Return%' then 'Return'
        	 when Status like '%RETURN%' then 'Return'
        	 when Status like '%Closed%' then 'Closed'
             else 'Others' end as Status
        into #SingleVisit
        from [treasure].[skiptrace].[dbo].[TSB_Single_SV_CaseTb] 
        
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        SQLREPORT.no,
        SQLREPORT.案號,
        SQLREPORT.姓名,
        Bank = SQLREPORT.銀行別,
        SQLREPORT.大金,
        PrioritySV = SQLREPORT.優先別,
        SV_Type = SQLREPORT.外訪種類,
        SQLREPORT.prioritydate,
        SQLREPORT.zip,
        City = SQLREPORT.縣市,
        Town = SQLREPORT.鄉鎮市區,
        address = SQLREPORT.地址,
        SQLREPORT.經度,
        SQLREPORT.緯度,
        Memo = SQLREPORT.備註,
        SQLREPORT.Priority,
        Objective = SQLREPORT.目的,
        Motivation = SQLREPORT.動機,
        OC = SQLREPORT.外訪員,
        AA = SQLREPORT.案件承辦,
        AA_contact = SQLREPORT.案件承辦分機,
        SQLREPORT.[預計外訪日/支援法務執行日],
        SQLREPORT.外訪確認,
        SQLREPORT.外訪確認日,
        SQLREPORT.addressi,
        SQLREPORT.BranchType,
        SQLREPORT.外訪對象,
        case when SQLREPORT.緯度 is null then 
            ( case when court2.緯度 is null then court.緯度
                    else court2.緯度 end)
        	else SQLREPORT.緯度 end as Latitude_r,
        case when SQLREPORT.經度 is null then 
            ( case when court2.經度 is null then court.經度
                    else court2.經度 end)
        	else SQLREPORT.經度 end as Longitude_r
        	,
        case when 目的 like '%執行時間%' then 
            convert(varchar(10),SUBSTRING(目的,patindex('%執行時間%',目的)+5,patindex('%執行時間%',目的)+9) )
            else '' end as cesv_TIME 
        
        into #temp1
        from [treasure].[skiptrace].[dbo].[SVTb_PrioritySV_LIST] SQLREPORT
        left join #DELLIST DEL on SQLREPORT.no = DEL.case_sv_no
        left join  (select distinct 動機,緯度,經度
                   from LOCATION_COURT
                  where 動機 <>'') as court on SQLREPORT.動機=court.動機
        left join (select distinct 地址, 緯度, 經度
                   from LOCATION_COURT )as court2 on SQLREPORT.地址=court2.地址
        where DEL.case_sv_no is null and                                               /*排除已訪案件*/
              (SQLREPORT.地址 not like '%地號%' or court2.地址 is not null)             /*排除異常址(地號)案件*/
        
        
        /*=======因併入台新單委，重新調整外訪優先等級，並考量CESV案件分拆執行日為圖層日期即第二碼為A、非圖層日為B=======*/
        /*=======新增案件類型：UCS自購案件、原業務案件=======*/

        select distinct
        Data_Date,
        convert(varchar(80),no)as no,
        CaseI = convert(varchar(80),[案號]),
        CM_Name = [姓名],
        Bank ,
        Debt_AMT = [大金],
        case when PrioritySV='1' then 
                (case when convert(varchar(10),cesv_TIME,111)=Data_Date
        		  then '1A'  else '1B' end)
        	 when PrioritySV='4' then '6'
        	 when PrioritySV='5' then '7'
        	 when PrioritySV='6' then '8'
        	 when PrioritySV='7' then '9'
        	 when PrioritySV='8' then '10'
        	 when PrioritySV='9' then '11'
        	 else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        case when Bank like '%台灣之星%' then 'UCS自購案件'
             else '原業務案件' end as Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude_r as Latitude,
        Longitude_r as Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        SUBSTRING(OC,1,4) as ocempi,
        SUBSTRING(OC,6,20) as OC,
        AA,
        AA_contact,
        cesv_TIME,
        '' as countdown
        into #temp2
        from #temp1
        where Latitude_r <> '' and Latitude_r  <> 0
        -- drop table #temp2
        /*=====================TSB單委案件等級給定，並考量分拆排訪日為圖層日期即第二碼為A、非圖層日為B=====================*/
        /*=======計算案件截至到期日天數=======*/
        
        select 
        convert(varchar(10),getdate()+1,111) as Data_date,
        convert(varchar(80),no) as no,
        CaseI,
        CM_Name,
        Bank,
        Debt_AMT,
        case when PrioritySV='4' then 
                   (case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '4A'  else '4B' end)
             when PrioritySV='5' then 
                  ( case when cesv_TIME <> '' and cesv_TIME<=convert(varchar(10),getdate()+1,111)
        		    then '5A'  else '5B' end)
             else PrioritySV end as PrioritySV_Adj,
        SV_Type,
        Case_Type,
        prioritydate,
        ZIP,
        City,
        Town,
        Address,
        Latitude,
        Longitude,
        Memo,
        Priority,
        Objective,
        Motivation,
        ocempi,
        OC,
        AA,
        AA_contact,
        cesv_TIME,
        datediff(day,convert(varchar(10),getdate()+1,111),dateadd(day,1,cesv_TIME)) as  countdown
        into #TSB
        from #SingleVisit
        where status = ''
        
        select 
            CaseI,
            Debt_AMT,
            PrioritySV_Adj,
            Case_Type,
            Address,
            Latitude,
            Longitude,
            Motivation,
            AA,
            AA_contact,
            cesv_TIME 
        from #TSB where OC = 'WHITE5082'
    """    
    cursor.execute(script)
    c_src = cursor.fetchall()
    cursor.close()
    conn.close()
    return c_src


#ALEXY_All
excel = ALEXY_All(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','執行時間'])
result.to_excel(rf'\\fortune\UCS\DI\OC\ALEXY\ALEXY_All.xlsx',index=False)

#ALEXY_TSB
excel = ALEXY_TSB(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','countdown'])
result.to_excel(rf'\\fortune\UCS\DI\OC\ALEXY\ALEXY_TSB.xlsx',index=False)

#ANDERSON_All
excel = ANDERSON_All(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','執行時間'])
result.to_excel(rf'\\fortune\UCS\DI\OC\ANDERSON\ANDERSON_All.xlsx',index=False)

#ANDERSON_TSB
excel = ANDERSON_TSB(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','countdown'])
result.to_excel(rf'\\fortune\UCS\DI\OC\ANDERSON\ANDERSON_TSB.xlsx',index=False)

#BEN4423_All
excel = BEN4423_All(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','執行時間'])
result.to_excel(rf'\\fortune\UCS\DI\OC\BEN4423\BEN4423_All.xlsx',index=False)

#BEN4423_TSB
excel = BEN4423_TSB(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','countdown'])
result.to_excel(rf'\\fortune\UCS\DI\OC\BEN4423\BEN4423_TSB.xlsx',index=False)

#BOO5056_All
excel = BOO5056_All(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','執行時間'])
result.to_excel(rf'\\fortune\UCS\DI\OC\BOO5056\BOO5056_All.xlsx',index=False)

#BOO5056_TSB
excel = BOO5056_TSB(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','countdown'])
result.to_excel(rf'\\fortune\UCS\DI\OC\BOO5056\BOO5056_TSB.xlsx',index=False)


#JACKYH_All
excel = JACKYH_All(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','執行時間'])
result.to_excel(rf'\\fortune\UCS\DI\OC\JACKYH\JACKYH_All.xlsx',index=False)

#JACKYH_TSB
excel = JACKYH_TSB(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','countdown'])
result.to_excel(rf'\\fortune\UCS\DI\OC\JACKYH\JACKYH_TSB.xlsx',index=False)

#JASON4703_All
excel = JASON4703_All(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','執行時間'])
result.to_excel(rf'\\fortune\UCS\DI\OC\JASON4703\JASON4703_All.xlsx',index=False)

#JASON4703_TSB
excel = JASON4703_TSB(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','countdown'])
result.to_excel(rf'\\fortune\UCS\DI\OC\JASON4703\JASON4703_TSB.xlsx',index=False)

#JULIAN_All
excel = JULIAN_All(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','執行時間'])
result.to_excel(rf'\\fortune\UCS\DI\OC\JULIAN\JULIAN_All.xlsx',index=False)

#JULIAN_TSB
excel = JULIAN_TSB(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','countdown'])
result.to_excel(rf'\\fortune\UCS\DI\OC\JULIAN\JULIAN_TSB.xlsx',index=False)


#SCOTT4162_All
excel = SCOTT4162_All(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','執行時間'])
result.to_excel(rf'\\fortune\UCS\DI\OC\SCOTT4162\SCOTT4162_All.xlsx',index=False)

#SCOTT4162_TSB
excel = SCOTT4162_TSB(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','countdown'])
result.to_excel(rf'\\fortune\UCS\DI\OC\SCOTT4162\SCOTT4162_TSB.xlsx',index=False)


#SIMONSH_All
excel = SIMONSH_All(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','執行時間'])
result.to_excel(rf'\\fortune\UCS\DI\OC\SIMONSH\SIMONSH_All.xlsx',index=False)

#SIMONSH_TSB
excel = SIMONSH_TSB(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','countdown'])
result.to_excel(rf'\\fortune\UCS\DI\OC\SIMONSH\SIMONSH_TSB.xlsx',index=False)


#SPANELY_All
excel = SPANELY_All(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','執行時間'])
result.to_excel(rf'\\fortune\UCS\DI\OC\SPANELY\SPANELY_All.xlsx',index=False)

#SPANELY_TSB
excel = SPANELY_TSB(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','countdown'])
result.to_excel(rf'\\fortune\UCS\DI\OC\SPANELY\SPANELY_TSB.xlsx',index=False)


#WHITE5082_All
excel = WHITE5082_All(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','執行時間'])
result.to_excel(rf'\\fortune\UCS\DI\OC\WHITE5082\WHITE5082_All.xlsx',index=False)

#WHITE5082_TSB
excel = WHITE5082_TSB(server,username,password,database)
result = pd.DataFrame(excel,columns=['案號','欠款金額','優先別','案件類型','地址','緯度','經度','動機','案件承辦','案件承辦分機','countdown'])
result.to_excel(rf'\\fortune\UCS\DI\OC\WHITE5082\WHITE5082_TSB.xlsx',index=False)