
from asyncio.windows_events import NULL
from math import nan
import os,sys
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import numpy as np
import pandas as pd
from datetime import datetime, date,time

import xlwings as xw
import copy

#RICEF Dataのシート名
ORIGININPUTSHEET = 'RICEF Data'
#シート作成用のアレイ
sheet_arr = ['FD(Lead)', 'FD(SCT)','UTC-S(Lead)', 'UTC-S(SCT)', 'UTC-S(Conditions)', 
                'TD(Lead)', 'TD(SCT)', 'Code(Lead)', 'Code(SCT)', 
                'UTC-E(Lead)', 'UTC-E(SCT)', 'UTC-E(FDer)','UTC-E(Lead_Defect)', 'UTC-E(SCT_Defect)', 'UTC-E(FDer_Defect)',
                'UTT(Lead)', 'UTT(SCT)','AT(FD)(Lead)','AT(FD)(SCT)','AT(Build)(Lead)','AT(Build)(SCT)',
                'ST(FD)(Lead)','ST(FD)(SCT)','ST(Build)(Lead)','ST(Build)(SCT)','ICR(Lead)','ICR(SCT)']   
##シートソート順するため                   
sheet_arr_oder = ['FD(Lead)', 'FD(SCT)','FD(SHIFT)','FD(Client)','FD(SCT_Lead)','UTC-S(Lead)', 'UTC-S(SCT)','UTC-S(SHIFT)',
                'UTC-S(Client)','UTC-S(Conditions)','TD(Lead)', 'TD(SCT)','TD(Client)','TD(JQE)', 'Code(Lead)', 'Code(SCT)','Code(Client)', 
                'UTC-E(Lead)', 'UTC-E(SCT)', 'UTC-E(FDer)','UTC-E(SHIFT)','UTC-E(Client)','UTC-E(Lead_Defect)', 'UTC-E(SCT_Defect)',
                'UTC-E(FDer_Defect)','UTC-E(SHIFT_Defect)','UTC-E(Client_Defect)','UTC-E(Tester_Defect)','UTT(Lead)', 'UTT(SCT)',
                'UTT(SHIFT)','UTT(Client)','AT(FD)(Lead)','AT(FD)(SCT)','AT(Build)(Lead)','AT(Build)(SCT)',
                'ST(FD)(Lead)','ST(FD)(SCT)','ST(Build)(Lead)','ST(Build)(SCT)','ICR(Lead)','ICR(SCT)']                        

#dataframe作成用
outputdf = [{'Project':'sample',
            '領域':'sample',
            'RICEFタイプ':'sample',
            'RICEF':'sample',
            '開発拠点':'sample',
            'Name(JA)':'sample',
            'Complexity':'sample',
            'Phase':'sample',
            'Deliverables':'sample',
            '作業者':'sample',
            '作業対応日時':2022/1/1,
            'レビュー者':'sample',
            'レビュー対応日時':2022/1/1,
            'KPI':'sample',
            '集計対象':'X',
            '集計対象外':'X',
            'Reviewpoint':0,
            'Classification':'sample',
            'S-Low':nan,
            'Low':nan,
            'Nornal':nan,
            'High':nan,
            'S-High':nan,
            'SS-High':nan,
            'Status':nan,
            '基準値外カウント（全体）':nan,
            '基準値外カウント（アクション残）':nan,
            'Signal(Original)':2,
            'ProjectCount用':1,
            'sheetName':'sample'}]
classlist = ['S-Low','Low','Nornal','High','S-High','SS-High']
strangedf = [{'RICEF':'sample','Sheet':'sample','Blank':'sample'}]


###################################################################
## getphaseDeliverKPI()
## Argument:sheetName:output先のシート名
## Return:[Phase,Deliverable,KPI]
## Overview:各シートごとに対応するPhase,Deliverable,KPIを返す。
###################################################################
def getphaseDeliverKPI(sheetName):
    #列名ラベル一覧(KPI,Phase,Deliverable取得用)
    FDLead_label = ['Design','FD','# of Reviewpoint(Lead)']
    FDSCT_label = ['Design','FD','# of Reviewpoint(SCT)']
    FDSHIFT_label = ['Design','FD','# of Reviewpoint(SHIFT)']
    FDClient_label = ['Design','FD','# of Reviewpoint(Client)']
    FDSCT_Lead_label = ['Design','FD','# of Reviewpoint(SCT_Lead)']    
    UTCSLead_label = ['Design','UTC-S','# of Reviewpoint(Lead)']
    UTCSSCT_label = ['Design','UTC-S','# of Reviewpoint(SCT)']
    UTCSSHIFT_label = ['Design','UTC-S','# of Reviewpoint(SHIFT)']
    UTCSCon_label = ['Design','UTC-S','UTC-E  # of Conditions']
    UTCSClient_label = ['Design','UTC-S','# of Reviewpoint(Client)']    
    TDLead_label = ['Build','TD','# of Reviewpoint(Lead)']
    TDSCT_label = ['Build','TD','# of Reviewpoint(SCT)']
    TDClient_label = ['Build','TD','# of Reviewpoint(Client)']
    TDJQE_label = ['Build','TD','# of Reviewpoint(JQE)']
    CodeLead_label = ['Build','Code','# of Reviewpoint(Lead)']
    CodeSCT_label = ['Build','Code','# of Reviewpoint(SCT)']
    CodeClient_label = ['Build','Code','# of Reviewpoint(Client)']
    UTCELead_label = ['Build','UTC-E','# of Reviewpoint(Lead)']
    UTCESCT_label = ['Build','UTC-E','# of Reviewpoint(SCT)']
    UTCEFDer_label = ['Build','UTC-E','# of Reviewpoint(FDer)']
    UTCESHIFT_label = ['Build','UTC-E','# of Reviewpoint(SHIFT)']
    UTCEClient_label = ['Build','UTC-E','# of Reviewpoint(Client)']
    UTCELeaddef_label = ['Build','UTC-E','# of defect(Lead)']
    UTCESCTdef_label = ['Build','UTC-E','# of defect(SCT)']
    UTCEFDerdef_label = ['Build','UTC-E','# of defect(FDer)']
    UTCESHIFTdef_label = ['Build','UTC-E','# of defect(SHIFT)']
    UTCEClientdef_label = ['Build','UTC-E','# of defect(Client)']
    UTCETesterdef_label = ['Build','UTC-E','# of defect(Tester)']
    UTTLead_label = ['Build','UTT','# of Reviewpoint(Lead)']
    UTTSCT_label = ['Build','UTT','# of Reviewpoint(SCT)']
    UTTSHIFT_label = ['Build','UTT','# of Reviewpoint(SHIFT)']
    UTTClient_label = ['Build','UTT','# of Reviewpoint(Client)']
    ATFDLead_label = ['Test-AT','Defect','# of FD Defect']
    ATFDSCT_label = ['Test-AT','Defect','# of FD Defect']
    ATBuildLead_label = ['Test-AT','Defect','# of Build Defect']
    ATBuildSCT_label = ['Test-AT','Defect','# of Build Defect']
    STFDLead_label = ['Test-ST','Defect','# of FD Defect']
    STFDSCT_label = ['Test-ST','Defect','# of FD Defect']
    STBuildLead_label = ['Test-ST','Defect','# of Build Defect']
    STBuildSCT_label = ['Test-ST','Defect','# of Build Defect']
    ICRLead_label = ['Build','ICR','# of ICR']
    ICRSCT_label = ['Build','ICR','# of ICR']
    sample_label = ['Build','ICR','# of ICR']
    ###

    #列名ラベルとシート名の紐づけ
    label_dic = {
    'FD(Lead)':FDLead_label,
    'FD(SCT)':FDSCT_label,
    'FD(SHIFT)':FDSHIFT_label,
    'FD(Client)':FDClient_label,
    'FD(SCT_Lead)':FDSCT_Lead_label,   
    'UTC-S(Lead)':UTCSLead_label,
    'UTC-S(SCT)':UTCSSCT_label,
    'UTC-S(SHIFT)':UTCSSHIFT_label, 
    'UTC-S(Client)':UTCSClient_label, 
    'UTC-S(Conditions)':UTCSCon_label,
    'TD(Lead)':TDLead_label,
    'TD(SCT)':TDSCT_label,
    'TD(Client)':TDClient_label,
    'TD(JQE)':TDJQE_label,     
    'Code(Lead)':CodeLead_label,
    'Code(SCT)':CodeSCT_label,
    'Code(Client)':CodeClient_label, 
    'UTC-E(Lead)':UTCELead_label,
    'UTC-E(Lead_Defect)':UTCELeaddef_label,
    'UTC-E(SCT)':UTCESCT_label,
    'UTC-E(SCT_Defect)':UTCESCTdef_label,
    'UTC-E(FDer)':UTCEFDer_label,
    'UTC-E(FDer_Defect)':UTCEFDerdef_label,
    'UTC-E(SHIFT)':UTCESHIFT_label,
    'UTC-E(SHIFT_Defect)':UTCESHIFTdef_label,
    'UTC-E(Client)':UTCEClient_label,    
    'UTC-E(Client_Defect)':UTCEClientdef_label,
    'UTC-E(Tester_Defect)':UTCETesterdef_label,     
    'UTT(Lead)':UTTLead_label,
    'UTT(SCT)':UTTSCT_label,
    'UTT(SHIFT)':UTTSHIFT_label,
    'UTT(Client)':UTTClient_label, 
    'AT(Build)(Lead)':ATBuildLead_label,
    'AT(Build)(SCT)':ATBuildSCT_label,
    'AT(FD)(Lead)':ATFDLead_label,
    'AT(FD)(SCT)':ATFDSCT_label,
    'ST(Build)(Lead)':STBuildLead_label,
    'ST(Build)(SCT)':STBuildSCT_label,
    'ST(FD)(Lead)':STFDLead_label,
    'ST(FD)(SCT)':STFDSCT_label,
    'ICR(Lead)':ICRLead_label,
    'ICR(SCT)':ICRSCT_label,
    'sample':sample_label}

    return label_dic[sheetName]


###################################################################
## getOpeRev()
## Argument
##  sheetName   :output先のシート名
##  ricef       :RICEF番号
##  frame       :mainデータフレーム
##  rsidx       :「RESULT」と入力されているセルの列番号
##  prjName     :プロジェクト名
## Return:[対応者,対応日,レビュー者,レビュー日,Classification,Reviewpoint]
## Overview:ricef番号とシート名より、該当する対応者、対応日、レビュー者、レビュー日、Classification,Reviewpointを取得する。
###################################################################
def getOpeRev(sheetName,ricef,frame,rsidx,prjName):
    #Return用の入れ物を宣言。
    ope_rev_set = ['','','','','','']
    
    #シート名ごとの列名ラベル一覧（取得対象が存在しない場合はsampleに設定し、以下の処理で取得処理を飛ばしている。)
    #基本的に列名を指定しているが、Reviewpoint,Classificationに関しては列名が重複しているため列番号で指定している。
    FDLead_label = ['FD Creation PIC','FD Creation End(A)','FD Reviewer','FD Review End(A)',0,32]
    FDSCT_label = ['FD Creation PIC','FD Creation End(A)','FD SCT Reviewer','FD SCT Review End(A)',1,33]
    UTCSLead_label = ['UTC-S Creation PIC','UTC-S Creation End(A)','UTC-S Reviewer','UTC-S Review End(A)',2,43]
    UTCSSCT_label = ['UTC-S Creation PIC','UTC-S Creation End(A)','UTC-S SCT Reviewer','UTC-S Review End(A)',3,44]
    UTCSCon_label = ['UTC-S Creation PIC','UTC-S Creation End(A)','sample','sample',4,39]
    TDLead_label = ['TD Creation PIC','TD Creation End(A)','TD Reviewer','TD Review End(A)',5,53]
    TDSCT_label = ['TD Creation PIC','TD Creation End(A)','TD SCT Reviewer','TD SCT Review End(A)',6,54]
    CodeLead_label = ['Code Creation PIC','Code Creation End(A)','Code Reviewer','Code Review End(A)',7,63]
    CodeSCT_label = ['Code Creation PIC','Code Creation End(A)',60,'Code SCT Review End(A)',8,64]#謎の60について：Code SCT ReviewerはPRJによってTD SCT Reviewerになっていることもあるため、列番号で指定している。
    UTCELead_label = ['UTC-E PIC','UTC-E End(A)','UTC-E Reviewer','UTC-E Review End(A)',9,77]
    UTCESCT_label = ['UTC-E PIC','UTC-E End(A)','UTC-E SCT Reviewer','UTC-E SCT Review End(A)',10,78]
    UTCEFDer_label = ['UTC-E PIC','UTC-E End(A)','FD Creation PIC','UTC-E FDer Review End(A)',11,79]
    UTCELeaddef_label = ['UTC-E PIC','UTC-E End(A)','UTC-E Reviewer','UTC-E Review End(A)',12,80]
    UTCESCTdef_label = ['UTC-E PIC','UTC-E End(A)','UTC-E SCT Reviewer','UTC-E SCT Review End(A)',13,81]
    UTCEFDerdef_label = ['UTC-E PIC','UTC-E End(A)','FD Creation PIC','UTC-E FDer Review End(A)',14,82]
    UTTLead_label = ['UT Tech PIC','UT Tech End(A)','UT Tech Reviewer','UT Tech Review End(A)',15,91]
    UTTSCT_label = ['UT Tech PIC','UT Tech End(A)','UT Tech SCT Reviewer','UT Tech SCT Review Start(A)',16,92]
    ATFDLead_label = ['FD Creation PIC','sample','FD Reviewer','sample',18,94]
    ATFDSCT_label = ['FD Creation PIC','sample','FD SCT Reviewer','sample',18,94]
    ATBuildLead_label = ['Code Creation PIC','sample','Code Reviewer','sample',19,95]
    ATBuildSCT_label = ['Code Creation PIC','sample',60,'sample',19,95]
    STFDLead_label = ['FD Creation PIC','sample','FD Reviewer','sample',20,96]
    STFDSCT_label = ['FD Creation PIC','sample','FD SCT Reviewer','sample',20,96]
    STBuildLead_label = ['Code Creation PIC','sample','Code Reviewer','sample',21,97]
    STBuildSCT_label = ['Code Creation PIC','sample',60,'sample',21,97]
    ICRLead_label  = ['FD Creation PIC','sample','FD Reviewer','FD Review End(A)',17,93]
    ICRSCT_label = ['FD Creation PIC','sample','FD SCT Reviewer','FD SCT Review End(A)',17,93]
    sample_label = ['sample','sample','sample','sample',17,93]
    
    ##ここから下はダミー
    FDSHIFT_label = ['sample','sample','sample','sample',17,93]
    FDClient_label = ['sample','sample','sample','sample',17,93]
    FDSCT_Lead_label = ['sample','sample','sample','sample',17,93]
    UTCSSHIFT_label =['sample','sample','sample','sample',17,93]
    UTCESHIFT_label = ['sample','sample','sample','sample',17,93]
    UTCESHIFTdef_label = ['sample','sample','sample','sample',17,93]
    UTTSHIFT_label = ['sample','sample','sample','sample',17,93]
    UTCSClient_label =['sample','sample','sample','sample',17,93]
    TDClient_label = ['sample','sample','sample','sample',17,93]
    TDJQE_label = ['sample','sample','sample','sample',17,93]    
    CodeClient_label = ['sample','sample','sample','sample',17,93]
    UTCEClient_label = ['sample','sample','sample','sample',17,93]
    UTCEClientdef_label = ['sample','sample','sample','sample',17,93]
    UTCETesterdef_label = ['sample','sample','sample','sample',17,93]
    UTTClient_label = ['sample','sample','sample','sample',17,93]
    ###

    if prjName == 'NHSTEP2':
        FDLead_label = ['FD Creation PIC','FD Creation End(A)','FD Reviewer','FD Review End(A)',0,39]
        FDSCT_label = ['FD Creation PIC','FD Creation End(A)','FD SCT Reviewer','FD SCT Review End(A)',1,40]
        FDSHIFT_label = ['FD Creation PIC','FD Creation End(A)','FD SHIFT Reviewer','FD SHIFT Review End(A)',2,41]
        UTCSLead_label = ['UTC-S Creation PIC','UTC-S Creation End(A)','UTC-S Reviewer','UTC-S Review End(A)',3,51]
        UTCSSCT_label = ['UTC-S Creation PIC','UTC-S Creation End(A)','UTC-S SCT Reviewer','UTC-S Review End(A)',4,52]
        UTCSCon_label = ['UTC-S Creation PIC','UTC-S Creation End(A)','sample','sample',5,47]
        TDLead_label = ['TD Creation PIC','TD Creation End(A)','TD Reviewer','TD Review End(A)',6,63]
        TDJQE_label = ['TD Creation PIC','TD Creation End(A)','TD JQE Reviewer','TD JQE Review End(A)',7,64]
        TDSCT_label = ['TD Creation PIC','TD Creation End(A)','TD SCT Reviewer','TD SCT Review End(A)',8,65]
        CodeLead_label = ['Code Creation PIC','Code Creation End(A)','Code Reviewer','Code Review End(A)',9,74]
        CodeSCT_label = ['Code Creation PIC','Code Creation End(A)',71,'Code SCT Review End(A)',10,75]#謎の60について：Code SCT ReviewerはPRJによってTD SCT Reviewerになっていることもあるため、列番号で指定している。
        UTCELead_label = ['UTC-E PIC','UTC-E End(A)','UTC-E Reviewer','UTC-E Review End(A)',11,90]
        UTCESCT_label = ['UTC-E PIC','UTC-E End(A)','UTC-E SCT Reviewer','UTC-E SCT Review End(A)',12,91]
        UTCEFDer_label = ['UTC-E PIC','UTC-E End(A)','FD Creation PIC','UTC-E FDer Review End(A)',13,92]
        UTCESHIFT_label = ['UTC-E PIC','UTC-E End(A)','UTC-E SHIFT Reviewer','UTC-E SHIFT Review End(A)',14,93]
        UTCELeaddef_label = ['UTC-E PIC','UTC-E End(A)','UTC-E Reviewer','UTC-E Review End(A)',15,94]
        UTCESCTdef_label = ['UTC-E PIC','UTC-E End(A)','UTC-E SCT Reviewer','UTC-E SCT Review End(A)',16,95]
        UTCEFDerdef_label = ['UTC-E PIC','UTC-E End(A)','FD Creation PIC','UTC-E FDer Review End(A)',17,96]
        UTCESHIFTdef_label = ['UTC-E PIC','UTC-E End(A)','UTC-E SHIFT Reviewer','UTC-E SHIFT Review End(A)',18,97]
        UTTLead_label = ['UT Tech PIC','UT Tech End(A)','UT Tech Reviewer','UT Tech Review End(A)',19,109]
        UTTSCT_label = ['UT Tech PIC','UT Tech End(A)','UT Tech SCT Reviewer','UT Tech SCT Review End(A)',20,110]
        UTTSHIFT_label = ['UT Tech PIC','UT Tech End(A)','UT Tech SHIFT Reviewer','UT Tech SHIFT Review End(A)',21,111]
        ATFDLead_label = ['FD Creation PIC','sample','FD Reviewer','sample',23,113]
        ATFDSCT_label = ['FD Creation PIC','sample','FD SCT Reviewer','sample',23,113]
        ATBuildLead_label = ['Code Creation PIC','sample','Code Reviewer','sample',24,114]
        ATBuildSCT_label = ['Code Creation PIC','sample',71,'sample',24,114]
        STFDLead_label = ['FD Creation PIC','sample','FD Reviewer','sample',25,115]
        STFDSCT_label = ['FD Creation PIC','sample','FD SCT Reviewer','sample',25,115]
        STBuildLead_label = ['Code Creation PIC','sample','Code Reviewer','sample',26,116]
        STBuildSCT_label = ['Code Creation PIC','sample',71,'sample',26,116]
        ICRLead_label  = ['FD Creation PIC','sample','FD Reviewer','FD Review End(A)',22,112]
        ICRSCT_label = ['FD Creation PIC','sample','FD SCT Reviewer','FD SCT Review End(A)',22,112]
        sample_label = ['sample','sample','sample','sample',22,112]
    ###
    elif prjName == 'SC':
        FDLead_label = ['FD Creation PIC','FD Creation End(A)','FD Reviewer','FD Review End(A)',0,41]
        FDSCT_label = ['FD Creation PIC','FD Creation End(A)','FD SCT Reviewer','FD SCT Review End(A)',1,42]
        FDClient_label = ['FD Creation PIC','FD Creation End(A)','FD Client Reviewer','FD Client Review End(A)',2,43]
        UTCSLead_label = ['UTC-S Creation PIC','UTC-S Creation End(A)','UTC-S Reviewer','UTC-S Review End(A)',3,55]
        UTCSSCT_label = ['UTC-S Creation PIC','UTC-S Creation End(A)','UTC-S SCT Reviewer','UTC-S Review End(A)',4,56]
        UTCSCon_label = ['UTC-S Creation PIC','UTC-S Creation End(A)','sample','sample',5,51]
        UTCSClient_label = ['UTC-S Creation PIC','UTC-S Creation End(A)','UTC-S Client Reviewer','UTC-S Client Review End(A)',6,57]
        TDLead_label = ['TD Creation PIC','TD Creation End(A)','TD Reviewer','TD Review End(A)',7,68]
        TDSCT_label = ['TD Creation PIC','TD Creation End(A)','TD SCT Reviewer','TD SCT Review End(A)',8,69]
        TDClient_label = ['TD Creation PIC','TD Creation End(A)','TD Client Reviewer','TD Client Review End(A)',9,70]
        CodeLead_label = ['Code Creation PIC','Code Creation End(A)','Code Reviewer','Code Review End(A)',10,81]
        CodeSCT_label = ['Code Creation PIC','Code Creation End(A)',77,'Code SCT Review End(A)',11,82]#謎の60について：Code SCT ReviewerはPRJによってTD SCT Reviewerになっていることもあるため、列番号で指定している。
        CodeClient_label = ['Code Creation PIC','Code Creation End(A)','Code Client Reviewer','Code Client Review End(A)',12,83]
        UTCELead_label = ['UTC-E PIC','UTC-E End(A)','UTC-E Reviewer','UTC-E Review End(A)',13,98]
        UTCESCT_label = ['UTC-E PIC','UTC-E End(A)','UTC-E SCT Reviewer','UTC-E SCT Review End(A)',14,99]
        UTCEFDer_label = ['UTC-E PIC','UTC-E End(A)','FD Creation PIC','UTC-E FDer Review End(A)',15,100]
        UTCEClient_label = ['UTC-E PIC','UTC-E End(A)','UTC-E Client Reviewer','UTC-E Client Review End(A)',16,101]
        UTCELeaddef_label = ['UTC-E PIC','UTC-E End(A)','UTC-E Reviewer','UTC-E Review End(A)',17,102]
        UTCESCTdef_label = ['UTC-E PIC','UTC-E End(A)','UTC-E SCT Reviewer','UTC-E SCT Review End(A)',18,103]
        UTCEFDerdef_label = ['UTC-E PIC','UTC-E End(A)','FD Creation PIC','UTC-E FDer Review End(A)',19,104]
        UTCEClientdef_label = ['UTC-E PIC','UTC-E End(A)','UTC-E Client Reviewer','UTC-E Client Review End(A)',20,105]
        UTTLead_label = ['UT Tech PIC','UT Tech End(A)','UT Tech Reviewer','UT Tech Review End(A)',21,116]
        UTTSCT_label = ['UT Tech PIC','UT Tech End(A)','UT Tech SCT Reviewer','UT Tech SCT Review Start(A)',22,117]
        UTTClient_label = ['UT Tech PIC','UT Tech End(A)','UT Tech Client Reviewer','UT Tech Client Review End(A)',23,118]
        ATFDLead_label = ['FD Creation PIC','sample','FD Reviewer','sample',25,120]
        ATFDSCT_label = ['FD Creation PIC','sample','FD SCT Reviewer','sample',25,120]
        ATBuildLead_label = ['Code Creation PIC','sample','Code Reviewer','sample',26,121]
        ATBuildSCT_label = ['Code Creation PIC','sample',77,'sample',26,121]
        STFDLead_label = ['FD Creation PIC','sample','FD Reviewer','sample',27,122]
        STFDSCT_label = ['FD Creation PIC','sample','FD SCT Reviewer','sample',27,122]
        STBuildLead_label = ['Code Creation PIC','sample','Code Reviewer','sample',28,123]
        STBuildSCT_label = ['Code Creation PIC','sample',77,'sample',28,123]
        ICRLead_label  = ['FD Creation PIC','sample','FD Reviewer','FD Review End(A)',24,119]
        ICRSCT_label = ['FD Creation PIC','sample','FD SCT Reviewer','FD SCT Review End(A)',24,119]
        sample_label = ['sample','sample','sample','sample',24,119]
    ###
    elif prjName == 'SQEX' or  prjName == 'Nintendo':
        FDLead_label = ['FD Creation PIC','FD Creation End(A)','FD Reviewer','FD Review End(A)',0,40]
        FDSCT_label = ['FD Creation PIC','FD Creation End(A)','FD SCT Reviewer','FD SCT Review End(A)',1,41]
        FDSHIFT_label = ['FD Creation PIC','FD Creation End(A)','FD SHIFT Reviewer','FD SHIFT Review End(A)',2,42]
        UTCSLead_label = ['UTC-S Creation PIC','UTC-S Creation End(A)','UTC-S Reviewer','UTC-S Review End(A)',3,54]
        UTCSSCT_label = ['UTC-S Creation PIC','UTC-S Creation End(A)','UTC-S SCT Reviewer','UTC-S Review End(A)',4,55]
        UTCSSHIFT_label = ['UTC-S Creation PIC','UTC-S Creation End(A)','UTC-S SHIFT Reviewer','UTC-S SHIFT Review End(A)',5,56]
        UTCSCon_label = ['UTC-S Creation PIC','UTC-S Creation End(A)','sample','sample',6,50]
        TDLead_label = ['TD Creation PIC','TD Creation End(A)','TD Reviewer','TD Review End(A)',7,65]
        TDSCT_label = ['TD Creation PIC','TD Creation End(A)','TD SCT Reviewer','TD SCT Review End(A)',8,66]
        CodeLead_label = ['Code Creation PIC','Code Creation End(A)','Code Reviewer','Code Review End(A)',9,75]
        CodeSCT_label = ['Code Creation PIC','Code Creation End(A)',72,'Code SCT Review End(A)',10,76]#謎の60について：Code SCT ReviewerはPRJによってTD SCT Reviewerになっていることもあるため、列番号で指定している。
        UTCELead_label = ['UTC-E PIC','UTC-E End(A)','UTC-E Reviewer','UTC-E Review End(A)',11,91]
        UTCESCT_label = ['UTC-E PIC','UTC-E End(A)','UTC-E SCT Reviewer','UTC-E SCT Review End(A)',12,92]
        UTCEFDer_label = ['UTC-E PIC','UTC-E End(A)','FD Creation PIC','UTC-E FDer Review End(A)',13,93]
        UTCESHIFT_label = ['UTC-E PIC','UTC-E End(A)','UTC-E SHIFT Reviewer','UTC-E SHIFT Review End(A)',14,94]
        UTCELeaddef_label = ['UTC-E PIC','UTC-E End(A)','UTC-E Reviewer','UTC-E Review End(A)',15,96]
        UTCESCTdef_label = ['UTC-E PIC','UTC-E End(A)','UTC-E SCT Reviewer','UTC-E SCT Review End(A)',16,97]
        UTCEFDerdef_label = ['UTC-E PIC','UTC-E End(A)','FD Creation PIC','UTC-E FDer Review End(A)',17,98]
        UTCESHIFTdef_label = ['UTC-E PIC','UTC-E End(A)','UTC-E SHIFT Reviewer','UTC-E SHIFT Review End(A)',18,99]
        UTCETesterdef_label = ['UTC-E PIC','UTC-E End(A)','sample','sample',19,95]
        UTTLead_label = ['UT Tech PIC','UT Tech End(A)','UT Tech Reviewer','UT Tech Review End(A)',20,110]
        UTTSCT_label = ['UT Tech PIC','UT Tech End(A)','UT Tech SCT Reviewer','UT Tech SCT Review Start(A)',21,111]
        UTTSHIFT_label = ['UT Tech PIC','UT Tech End(A)','UT Tech SHIFT Reviewer','UT Tech SHIFT Review End(A)',22,112]
        ATFDLead_label = ['FD Creation PIC','sample','FD Reviewer','sample',24,114]
        ATFDSCT_label = ['FD Creation PIC','sample','FD SCT Reviewer','sample',24,114]
        ATBuildLead_label = ['Code Creation PIC','sample','Code Reviewer','sample',25,115]
        ATBuildSCT_label = ['Code Creation PIC','sample',72,'sample',25,115]
        STFDLead_label = ['FD Creation PIC','sample','FD Reviewer','sample',26,116]
        STFDSCT_label = ['FD Creation PIC','sample','FD SCT Reviewer','sample',26,116]
        STBuildLead_label = ['Code Creation PIC','sample','Code Reviewer','sample',27,117]
        STBuildSCT_label = ['Code Creation PIC','sample',72,'sample',27,117]
        ICRLead_label  = ['FD Creation PIC','sample','FD Reviewer','FD Review End(A)',23,113]
        ICRSCT_label = ['FD Creation PIC','sample','FD SCT Reviewer','FD SCT Review End(A)',23,113]
        sample_label = ['sample','sample','sample','sample',23,113]
 
    if prjName == 'SeikoEpson':
            FDLead_label = ['FD Creation PIC','FD Creation End(A)','FD Reviewer','FD Review End(A)',0,34]
            FDSCT_label = ['FD Creation PIC','FD Creation End(A)','FD SCT Reviewer','FD SCT Review End(A)',1,35]
            FDSCT_Lead_label = ['FD Creation PIC','FD Creation End(A)','sample','sample',2,36]
            UTCSLead_label = ['UTC-S Creation PIC','UTC-S Creation End(A)','UTC-S Reviewer','UTC-S Review End(A)',3,48]
            UTCSSCT_label = ['UTC-S Creation PIC','UTC-S Creation End(A)','UTC-S SCT Reviewer','UTC-S Review End(A)',4,49]
            UTCSSHIFT_label = ['UTC-S Creation PIC','UTC-S Creation End(A)','UTC-S SHIFT Reviewer','UTC-S SHIFT Review End(A)',5,50]
            UTCSCon_label = ['UTC-S Creation PIC','UTC-S Creation End(A)','sample','sample',6,44]
            TDLead_label = ['TD Creation PIC','TD Creation End(A)','TD Reviewer','TD Review End(A)',7,59]
            TDSCT_label = ['TD Creation PIC','TD Creation End(A)','TD SCT Reviewer','TD SCT Review End(A)',8,60]
            CodeLead_label = ['Code Creation PIC','Code Creation End(A)','Code Reviewer','Code Review End(A)',9,69]
            CodeSCT_label = ['Code Creation PIC','Code Creation End(A)',66,'Code SCT Review End(A)',10,70]#謎の60について：Code SCT ReviewerはPRJによってTD SCT Reviewerになっていることもあるため、列番号で指定している。
            UTCELead_label = ['UTC-E PIC','UTC-E End(A)','UTC-E Reviewer','UTC-E Review End(A)',11,83]
            UTCESCT_label = ['UTC-E PIC','UTC-E End(A)','UTC-E SCT Reviewer','UTC-E SCT Review End(A)',12,84]
            UTCEFDer_label = ['UTC-E PIC','UTC-E End(A)','FD Creation PIC','UTC-E FDer Review End(A)',13,85]
            UTCELeaddef_label = ['UTC-E PIC','UTC-E End(A)','UTC-E Reviewer','UTC-E Review End(A)',14,86]
            UTCESCTdef_label = ['UTC-E PIC','UTC-E End(A)','UTC-E SCT Reviewer','UTC-E SCT Review End(A)',15,87]
            UTCEFDerdef_label = ['UTC-E PIC','UTC-E End(A)','FD Creation PIC','UTC-E FDer Review End(A)',16,88]
            UTTLead_label = ['UT Tech PIC','UT Tech End(A)','UT Tech Reviewer','UT Tech Review End(A)',17,97]
            UTTSCT_label = ['UT Tech PIC','UT Tech End(A)','UT Tech SCT Reviewer','UT Tech SCT Review Start(A)',18,98]
            ATFDLead_label = ['FD Creation PIC','sample','FD Reviewer','sample',20,100]
            ATFDSCT_label = ['FD Creation PIC','sample','FD SCT Reviewer','sample',20,100]
            ATBuildLead_label = ['Code Creation PIC','sample','Code Reviewer','sample',21,101]
            ATBuildSCT_label = ['Code Creation PIC','sample',66,'sample',21,101]
            STFDLead_label = ['FD Creation PIC','sample','FD Reviewer','sample',22,102]
            STFDSCT_label = ['FD Creation PIC','sample','FD SCT Reviewer','sample',22,102]
            STBuildLead_label = ['Code Creation PIC','sample','Code Reviewer','sample',23,103]
            STBuildSCT_label = ['Code Creation PIC','sample',66,'sample',23,103]
            ICRLead_label  = ['FD Creation PIC','sample','FD Reviewer','FD Review End(A)',19,99]
            ICRSCT_label = ['FD Creation PIC','sample','FD SCT Reviewer','FD SCT Review End(A)',19,99]
            sample_label = ['sample','sample','sample','sample',19,99]

    #列名ラベルとシート名の紐づけ
    label_dic = {
    'FD(Lead)':FDLead_label,
    'FD(SCT)':FDSCT_label,
    'FD(SHIFT)':FDSHIFT_label,
    'FD(Client)':FDClient_label,
    'FD(SCT_Lead)':FDSCT_Lead_label,
    'UTC-S(Lead)':UTCSLead_label,
    'UTC-S(SCT)':UTCSSCT_label,
    'UTC-S(SHIFT)':UTCSSHIFT_label,     
    'UTC-S(Client)':UTCSClient_label, 
    'UTC-S(Conditions)':UTCSCon_label,
    'TD(Lead)':TDLead_label,
    'TD(SCT)':TDSCT_label,
    'TD(Client)':TDClient_label,
    'TD(JQE)':TDJQE_label,
    'Code(Lead)':CodeLead_label,
    'Code(SCT)':CodeSCT_label,
    'Code(Client)':CodeClient_label,
    'UTC-E(Lead)':UTCELead_label,
    'UTC-E(Lead_Defect)':UTCELeaddef_label,
    'UTC-E(SCT)':UTCESCT_label,
    'UTC-E(SCT_Defect)':UTCESCTdef_label,
    'UTC-E(FDer)':UTCEFDer_label,
    'UTC-E(FDer_Defect)':UTCEFDerdef_label,
    'UTC-E(SHIFT)':UTCESHIFT_label,
    'UTC-E(SHIFT_Defect)':UTCESHIFTdef_label,
    'UTC-E(Client)':UTCEClient_label,
    'UTC-E(Client_Defect)':UTCEClientdef_label,
    'UTC-E(Tester_Defect)':UTCETesterdef_label,   
    'UTT(Lead)':UTTLead_label,
    'UTT(SCT)':UTTSCT_label,
    'UTT(SHIFT)':UTTSHIFT_label,
    'UTT(Client)':UTTClient_label,
    'AT(Build)(Lead)':ATBuildLead_label,
    'AT(Build)(SCT)':ATBuildSCT_label,
    'AT(FD)(Lead)':ATFDLead_label,
    'AT(FD)(SCT)':ATFDSCT_label,
    'ST(Build)(Lead)':STBuildLead_label,
    'ST(Build)(SCT)':STBuildSCT_label,
    'ST(FD)(Lead)':STFDLead_label,
    'ST(FD)(SCT)':STFDSCT_label,
    'ICR(Lead)':ICRLead_label,
    'ICR(SCT)':ICRSCT_label,
    'sample':sample_label}

    #シート名より列名ラベルの取得
    label = label_dic[sheetName]
    factornum = len(label)

    for i in range(0,factornum):
        colName = label[i]
        if colName == 'sample':
            continue
        #ラベルの中身が数値の場合は列番号より値を取得。
        if(type(colName) == int):
            colName = frame.columns[colName + rsidx]
        frid = frame.index[(frame['RICEF'] == ricef)]

        #対応日、レビュー対応日を取得する場合はdatetimeからdateに変換(時間情報は不要のため)
        if (i == 1) or (i == 3):
            if type(frame.loc[frid,colName].iloc[-1]) != str:
                if (type(frame.loc[frid,colName].iloc[-1]) == time) or (type(frame.loc[frid,colName].iloc[-1]) == int):
                    ope_rev_set[i] = nan
                else:
                    ope_rev_set[i] = frame.loc[frid,colName].iloc[-1].date()
            else:
                ope_rev_set[i] = frame.loc[frid,colName].iloc[-1]
        #対応者、レビュー対応者を取得する場合は名前に含まれる(),矢印等の除去処理を実施する。
        elif(i == 0) or (i == 2):
            name = frame.loc[frid,colName].iloc[-1]
            if type(name) == str:
                name = nameModify(name)
                #半角スペースが入っている場合があるので空白に変換する。
                if name == ' ':
                    name = ''
            ope_rev_set[i] = name
        else:
            ope_rev_set[i] = frame.loc[frid,colName].iloc[-1]
    return ope_rev_set

###################################################################
## nameModify()
## Argument
##  name        :変換したい名前
## Return:変換後の名前
## Overview:下記のルールで名前を変換する。
##      1.（）,()が混ざっている際、カッコの中身が1文字より大きければ不要とみなす。
##      2.→,⇒が含まれている際は一番右側の矢印が指している名前を担当者名とする
###################################################################
def nameModify(name):
    if ('→' in name):
        if('（' in name):
            count = name.find('）') - name.find('（') 
            if count > 2:
                name = name[name.rfind('→') + 1:name.find('（') ]
            else:
                name = name[name.rfind('→') + 1:len(name)]
        elif('(' in name):
            count = name.find(')') - name.find('(')
            if count > 2:
                name = name[name.rfind('→') + 1:name.find('(') ]
            else:
                name = name[name.rfind('→') + 1:len(name)]
        else:
            name = name[name.rfind('→') + 1:len(name)]
    elif('⇒' in name):
        if('（' in name):
            count = name.find('）') - name.find('（') 
            if count > 2:
                name = name[name.rfind('⇒') + 1:name.find('（') ]
            else:
                name = name[name.rfind('⇒') + 1:len(name)]
        elif('(' in name):
            count = name.find(')') - name.find('(')
            if count > 2:
                name = name[name.rfind('⇒') + 1:name.find('(') ]
            else:
                name = name[name.rfind('⇒') + 1:len(name)]
        else:
            name = name[name.rfind('⇒') + 1:len(name)]
    else:
        if('（' in name):
            count = name.find('）') - name.find('（') 
            if count > 2:
                name = name[name.rfind('⇒') + 1:name.find('（') ]
            else:
                name = name[name.rfind('⇒') + 1:len(name)]
        elif('(' in name):
            count = name.find(')') - name.find('(')
            if count > 2:
                name = name[name.rfind('⇒') + 1:name.find('(') ]
            else:
                name = name[name.rfind('⇒') + 1:len(name)]
        else:
            name = name[name.rfind('⇒') + 1:len(name)]
    return name


###################################################################
## fileMake()
## Argument
##  outputoath      :output先フォルダ
##  mainfilepath    :読み込みファイル。        
## Return:アウトプット先のファイル名
## Overview:
##      1.インポートしたエクセルファイルからpowerbiにインポートするためのファイルを作成する。
###################################################################
def fileMake(outputpath,mainfilepath):
    #プロジェクト名作成を読み込みファイル名より取得。
    hanteidf = pd.DataFrame(strangedf)
    #例：ファイル名が「SCT_DEV_Quality_Reporting_MISUMI_20211117_v0.04_0510.xlsx]の場合、MISUMIを取得。(Reportingという文字列を基準に取得している。)
    prjName = mainfilepath[mainfilepath.find('Reporting_') + len('Reporting_'):mainfilepath.find('_',mainfilepath.find('Reporting_') + len('Reporting_'))]
    if prjName == 'NHSTEP2':
        print('NHプロジェクトのためSHIFTとTD(JQE)を考慮に入れます。')
        sheet_arr.append('FD(SHIFT)')
        sheet_arr.append('TD(JQE)')
        sheet_arr.append('UTC-E(SHIFT)')
        sheet_arr.append('UTC-E(SHIFT_Defect)')
        sheet_arr.append('UTT(SHIFT)')
    elif prjName == 'SC':
        print('SCプロジェクトのためClientを考慮に入れます。')
        sheet_arr.append('FD(Client)')
        sheet_arr.append('UTC-S(Client)')
        sheet_arr.append('TD(Client)')
        sheet_arr.append('Code(Client)')
        sheet_arr.append('UTC-E(Client)')
        sheet_arr.append('UTC-E(Client_Defect)')
        sheet_arr.append('UTT(Client)')
    elif prjName == 'SQEX' or  prjName == 'Nintendo':
        print('SQEX,NintendoプロジェクトのためSHIFT,Testerを考慮に入れます。')
        sheet_arr.append('FD(SHIFT)')
        sheet_arr.append('UTC-S(SHIFT)')
        sheet_arr.append('UTC-E(SHIFT)')
        sheet_arr.append('UTC-E(SHIFT_Defect)')
        sheet_arr.append('UTC-E(Tester_Defect)') 
        sheet_arr.append('UTT(SHIFT)')
    elif prjName == 'SeikoEpson':
        print('SeikoEpsonプロジェクトのためSCT_Lead,SHIFTを考慮に入れます。')
        sheet_arr.append('FD(SCT_Lead)')
        sheet_arr.append('UTC-S(SHIFT)')

    origindf = pd.DataFrame(outputdf)

    ##Resultと記載されている列のインデックス取得(この列を基準にReviewPoint等を取得する。)
    Rsidx = pd.read_excel(mainfilepath,sheet_name = ORIGININPUTSHEET,header = 1).columns.get_loc('RESULT')

    #インポートファイルの読み込み
    readdf = pd.read_excel(mainfilepath,sheet_name = ORIGININPUTSHEET,header = 3).fillna('').drop(0)
    
    #SQEXの場合は「開発担当拠点」、「FD集計対象外」、「UTC-S集計対象外」列があるのでsqedfとして避けておく。他のprjと動きを統一するためreaddfからは削除。
    if prjName == 'SQEX':
        sqedf = readdf[['RICEF','開発担当拠点']]
        fdcntout = readdf[['RICEF','FD集計対象外']]
        utcscntout = readdf[['RICEF','UTC-S集計対象外']]
        del readdf['開発担当拠点'],readdf['FD集計対象外'],readdf['UTC-S集計対象外']

    #deepdiveの読み込み
    deepdivedf = pd.read_excel(mainfilepath,sheet_name = 'Deep Dive Items')

    #RICEF一覧を取得(行終わりのeも取れてしまうのが玉に瑕)
    riceflist = readdf['RICEF'].unique()
    
    #RICEFを読み込んだ行分繰り返し
    for j in range(0,len(riceflist)):
        ricefName = riceflist[j]
        if (ricefName != 'e') and (ricefName != '')  :
            #maindf中の該当RICEFの行番号を取得
            ricefidx = readdf.index[(readdf['RICEF'] == ricefName)]
            #シート名分繰り返す。
            for i in range(0,len(sheet_arr)):
                #origindfに追加するためのdf一行分を宣言
                tempdf = copy.copy(outputdf)
                
                #既に取得しているProject,RICEFをtempdfに格納
                tempdf[0]['Project'] = prjName
                tempdf[0]['RICEF'] = ricefName

                #領域名の取得。(SeikoEpson以外はRICEFの頭2文字を領域名とする。)
                if (prjName == 'SeikoEpson') or (prjName == 'Toshiba'):
                    tempdf[0]['領域'] = readdf.loc[ricefidx,'Sub-Area'].iloc[-1]
                ###2022/07/14 add 7andI要件を追加　star
                #領域名の取得。(7andIの場合はAreaを領域名とする。)    
                elif prjName == '7andI':
                    tempdf[0]['領域'] = readdf.loc[ricefidx,'Area'].iloc[-1]    
                ###2022/07/14 add 7andI要件を追加　end    
                else:
                    tempdf[0]['領域'] = ricefName[0:2]
                tempdf[0]['RICEFタイプ'] = readdf.loc[ricefidx,'RICEF Type'].iloc[-1]

                ##開発拠点、集計対象外についてはSQEXの場合は別フレーム(sqedf)として保存した'開発担当拠点'、'FD集計対象外'、'UTC-S集計対象外'列より取得する。
                if prjName == 'SQEX':
                    tempdf[0]['開発拠点'] = sqedf.loc[ricefidx,'開発担当拠点'].iloc[-1]
                    if 'FD(Lead)' == sheet_arr[i] or 'FD(SCT)' == sheet_arr[i] or 'FD(SHIFT)' == sheet_arr[i]:
                        tempdf[0]['集計対象外'] = fdcntout.loc[ricefidx,'FD集計対象外'].iloc[-1]
                    elif 'UTC-S(Lead)' == sheet_arr[i] or 'UTC-S(SCT)' == sheet_arr[i] or 'UTC-S(Conditions)' == sheet_arr[i] \
                        or 'UTC-S(SHIFT)' == sheet_arr[i]:
                        tempdf[0]['集計対象外'] = utcscntout.loc[ricefidx,'UTC-S集計対象外'].iloc[-1]
                    else:
                        tempdf[0]['集計対象外'] = nan
                else:
                    tempdf[0]['開発拠点'] = 'ATCI'#一旦SQEX以外はATCI固定にしておく

                #Name(JA)とComplexityはdfからそのまま取得。
                tempdf[0]['Name(JA)'] = readdf.loc[ricefidx,'Name(JA)'].iloc[-1]
                tempdf[0]['Complexity'] = readdf.loc[ricefidx,'Complexity'].iloc[-1]

                #getOpeRevにて作業者、作業対応日時、レビュー者、レビュー対応日時、Classification,Reviewpoint,sheetNameをまとめてoutlistに取得。
                outlist = getOpeRev(sheet_arr[i],ricefName,readdf,Rsidx,prjName)
                tempdf[0]['作業者'] = outlist[0]
                tempdf[0]['作業対応日時'] = outlist[1]
                tempdf[0]['レビュー者'] = outlist[2]
                tempdf[0]['レビュー対応日時'] = outlist[3]
                tempdf[0]['Classification'] = outlist[4]
                tempdf[0]['Reviewpoint'] = outlist[5]
                tempdf[0]['sheetName'] = sheet_arr[i]
                
                #getphaseDeliverKPIにてPhase,Deliverable,KPIをまとめて取得
                phasedelivKPIlist = getphaseDeliverKPI(sheet_arr[i])
                tempdf[0]['Phase'] = phasedelivKPIlist[0]
                tempdf[0]['Deliverables'] = phasedelivKPIlist[1]
                tempdf[0]['KPI'] = phasedelivKPIlist[2]

                #集計対象は基本的にXだが、なぜかUTC-Sのみ空白とする。（理由は知らない)
                if sheet_arr[i] == 'UTC-S(Lead)':
                    tempdf[0]['集計対象'] = nan
                else:
                    tempdf[0]['集計対象'] = 'X' 

                #Classificationに応じて該当列に1を入力する。
                classification = outlist[4]
                countflg = 0
                if classification == 'OK':
                    for k in range(0,len(classlist)):
                        tempdf[0][classlist[k]] = nan
                    tempdf[0]['Nornal'] = 1
                    tempdf[0]['基準値外カウント（全体）'] = 0
                else:
                    for k in range(0,len(classlist)):
                        if(classification == classlist[k]):
                            tempdf[0][classlist[k]] = 1
                            tempdf[0]['基準値外カウント（全体）'] = 1
                            countflg = 1
                        else:
                            tempdf[0][classlist[k]] = nan
                    if countflg == 0:
                        tempdf[0]['基準値外カウント（全体）'] = nan


                #tempdfをデータフレームに変換後、元のDFに追加
                adddf = pd.DataFrame(tempdf)
                origindf = pd.concat([origindf,adddf])

                
                #各行ごとにアカンかった(Reviewpointが入っているのに日付やClassificationが入っていないなど)部分をリスト化する。
                hanteidf = pd.concat([hanteidf,pd.DataFrame(hantei(tempdf))])


    #ouputファイルの作成処理。
    #「SCT_DEV_Quality_Report_[Project名]_yyyymmdd_.xlsx」
    dt = datetime.today()
    #outputfilename = 'SCT_DEV_Quality_Report_' + prjName + '_' + str(dt.year) + str(dt.month) + str(dt.day) + '.xlsx'
    outputfilename = 'SCT_DEV_Quality_Report_' + prjName + '_' + dt.strftime('%Y%m%d') + '.xlsx'

    #sheetName列を頼りに各シートにdataframeを展開していく。
    with pd.ExcelWriter(outputpath +'\\' + outputfilename ,datetime_format="YYYY/MM/DD") as writer:
        ##各シートを指定のソート順にする
        sheet_arr_new = sorted(sheet_arr, key=sheet_arr_oder.index)
        for sheetid in sheet_arr_new:
            outdf = origindf.query('sheetName == "' + sheetid + '"').drop('sheetName',axis = 1)
            ##SQEXの場合はFD,UTC-S以外のシートから集計対象外項目を削除
            if prjName == 'SQEX':
                if 'FD(Lead)' != sheetid and 'FD(SCT)' != sheetid and 'FD(SHIFT)' != sheetid and 'UTC-S(Lead)' != sheetid \
                    and 'UTC-S(SCT)' != sheetid and 'UTC-S(Conditions)' != sheetid and 'UTC-S(SHIFT)' != sheetid:
                    outdf.drop("集計対象外",axis=1,inplace=True)
            ##SQEX以外の場合は各シートの集計対象外項目を削除
            else:    
                outdf.drop("集計対象外",axis=1,inplace=True)
            ##ファイル書き出し    
            outdf.to_excel(writer,sheet_name = sheetid,index = False)
        #最後にdeepdiveItemsを追加する。
        deepdivedf.to_excel(writer,sheet_name= 'Deep Dive Items',index = False)

    #ブランク判定の結果をstrange.csvに吐き出す。    
    hanteidf.query('RICEF != "sample"').to_csv(outputpath + '\\' + 'blanklist_' + prjName + '_' + dt.strftime('%Y%m%d') \
        + '.csv',index = False,encoding='shift-jis')
    return outputfilename

###################################################################
## hantei()
## Argument
##  tempdf      :RICEF,Sheetごとの1行分のデータフレーム
## Return:hanteidf:RICEF,Sheet名、Blankの列名を返す。
## Overview:
##      1.Reviewpointが入力されていてかつレビュー日やClassificationがブランクであるかどうかを判定する。
###################################################################
def hantei(tempdf):
    hanteidf = copy.deepcopy(strangedf)
    conv = pd.DataFrame(tempdf)
    ###2022/07/13 chg Reviewpointがブランクで、Classificationが入力しているbug　star
    ##if type(conv.loc[0,'Reviewpoint']) == str:
    ##    conv.loc[0,'Reviewpoint'] = 'a'
    if not conv.loc[0,'Classification'] == '':
        if conv.loc[0,'Reviewpoint'] == '':
            hanteidf[0]['RICEF'] = tempdf[0]['RICEF']
            hanteidf[0]['Sheet'] = tempdf[0]['sheetName']
            hanteidf[0]['Blank'] = 'Reviewpoint'
            hanteidf[0]['Reviewpoint'] = tempdf[0]['Reviewpoint']
        elif (conv.loc[0,'レビュー対応日時'] == '') and (tempdf[0]['sheetName'] != 'UTC-S(Conditions)'):
            hanteidf[0]['RICEF'] = tempdf[0]['RICEF']
            hanteidf[0]['Sheet'] = tempdf[0]['sheetName']
            hanteidf[0]['Blank'] = 'レビュー日'    
    ##if not conv.loc[0,'Reviewpoint'] == 'a':
    if not conv.loc[0,'Reviewpoint'] == '':    
    ###2022/07/13 chg Reviewpointがブランクで、Classificationが入力しているbug    end
        if conv.loc[0,'Classification'] == '':
            hanteidf[0]['RICEF'] = tempdf[0]['RICEF']
            hanteidf[0]['Sheet'] = tempdf[0]['sheetName']
            hanteidf[0]['Blank'] = 'Classification'
    ###2022/07/13 chg Reviewpointがブランクで、Classificationが入力しているbug  star        
            hanteidf[0]['Classification'] = tempdf[0]['Classification']
    ###2022/07/13 chg Reviewpointがブランクで、Classificationが入力しているbug  end        
        elif (conv.loc[0,'レビュー対応日時'] == '') and (tempdf[0]['sheetName'] != 'UTC-S(Conditions)'):
            hanteidf[0]['RICEF'] = tempdf[0]['RICEF']
            hanteidf[0]['Sheet'] = tempdf[0]['sheetName']
            hanteidf[0]['Blank'] = 'レビュー日'
    return hanteidf

###################################################################
## ricefadd()
## Argument
##  outputpath      :output先フォルダ
##  riceffilepath   :読み込みファイル。RICEF Dataを持っているファイルのパス
#   outputfilename  :output先。つまりfileMakeで作成したファイルのパス        
## Return:
## Overview:
##      1.[RICEF Data]シートを書式を保持したままアウトプットファイルにコピー・追加する。(openpyxlでは無理なのでxlwingsを使うことに。)
###################################################################
def ricefadd(outputpath,riceffilePath,outputfilename):
    outFilepath = outputpath + '\\' + outputfilename
    bk_a = xw.Book(riceffilePath)
    bk_z = xw.Book(outFilepath)

    bk_asheet = bk_a.sheets['RICEF Data']
    sht = bk_z.sheets['Deep Dive Items']
    bk_asheet.copy(before=sht)

    bk_a.save()
    bk_z.save()
    bk_a.close()
    bk_z.close()
    return

# OUTPUTフォルダ指定の関数
def dirdialog_clicked():
    iDir = os.path.abspath(os.path.abspath(''))
    iDirPath = filedialog.askdirectory(initialdir = iDir)
    entry1.set(iDirPath)

# ファイル指定の関数
def filedialog_clicked():
    fTyp = [("エクセルファイル", ".xlsx")]
    iFile = os.path.abspath(os.path.abspath(''))
    iFilePath = filedialog.askopenfilename(filetype = fTyp, initialdir = iFile)
    entry2.set(iFilePath)

###################################################################
## conductMain()
## Argument     
## Return:
## Overview:
##      1.実行ボタンが押された際に起動。fileMake()とricefAdd()を呼ぶ。
###################################################################
def conductMain():
    outputPath = entry1.get()
    mainfilePath = entry2.get()

    #ファイル未指定の際の処理
    if (outputPath == '') or mainfilePath == '' or mainfilePath == '':
        errmsg = 'パスをすべて指定してください\nやり直しますか？'
        ret = messagebox.askyesno("確認", errmsg)
        if ret == False:
            messagebox.showinfo('info','終了します')
            myexit()
        else:
            return
    #try:
    outputfilename = fileMake(outputPath,mainfilePath)
    #except:
    #messagebox.showerror('err','Errorが発生したので終了します。')
    #myexit()
    ricefadd(outputPath,mainfilePath,outputfilename)
    messagebox.showinfo('info','ファイル出力完了しました')
    myexit()


# 謎の終了関数の定義(なぜだか知らないがexe化をした際にメイン処理のボタン押下時のcommandにてsys.exit()を認識しないことがあった。)
def myexit():
    sys.exit()

#メイン
if __name__ == "__main__":

    # rootの作成
    root = Tk()
    root.title("Importデータ作成")

    # Frame1の作成
    frame1 = ttk.Frame(root, padding=10)
    frame1.grid(row=0, column=1, sticky=E)

    # 「フォルダ参照」ラベルの作成
    IDirLabel = ttk.Label(frame1, text="Outputパス＞＞", padding=(5, 2))
    IDirLabel.pack(side=LEFT)

    # 「フォルダ参照」エントリーの作成
    entry1 = StringVar()
    IDirEntry = ttk.Entry(frame1, textvariable=entry1, width=30)
    IDirEntry.pack(side=LEFT)

    # 「フォルダ参照」ボタンの作成
    IDirButton = ttk.Button(frame1, text="参照", command=dirdialog_clicked)
    IDirButton.pack(side=LEFT)

    # Frame2の作成
    frame2 = ttk.Frame(root, padding=10)
    frame2.grid(row=2, column=1, sticky=E)

    # 「ファイル参照」ラベルの作成
    IFileLabel = ttk.Label(frame2, text="インプットファイル＞＞", padding=(5, 2))
    IFileLabel.pack(side=LEFT)

    # 「ファイル参照」エントリーの作成
    entry2 = StringVar()
    IFileEntry = ttk.Entry(frame2, textvariable=entry2, width=30)
    IFileEntry.pack(side=LEFT)

    # 「ファイル参照」ボタンの作成
    IFileButton = ttk.Button(frame2, text="参照", command=filedialog_clicked)
    IFileButton.pack(side=LEFT)

    # Frame4の作成
    frame4 = ttk.Frame(root, padding=10)
    frame4.grid(row=5,column=1,sticky=W)

    # 実行ボタンの設置
    button1 = ttk.Button(frame4, text="実行", command=conductMain)
    button1.pack(fill = "x", padx=60, side = "left")

    # キャンセルボタンの設置
    button2 = ttk.Button(frame4, text=("閉じる"), command=myexit)
    button2.pack(fill = "x", padx=60, side = "left")

    root.mainloop()