/**
 * Method to Turn an HTML Table/Grid to an excel document.  
 * This supports styles, currency, and date formats.  More can be added/adjusted, but this is just the basics.
 * @method      grid2Excel
 * @property    {string}        sLocalURI               This is the local browser URI to the base64 document that get's created
 * @property    {string}        sFinalFileXML           This is the template for the final XML document.  This is specifically for MSExcel
 * @property    {string}        sWorksheetTemplate      This is the worksheet template needed for excel to render.
 * @property    {string}        sCellTemplate           This is the data cell template for the excel sheet
 * @property    {function}      base64                  This is the base64 encoding method that creates the local url
 * @property    {function}      fnFormat                This method will merge the data with the template & properly format the data
 * @property    {function}      fnCreateWorkbooks       This method is the workbook creator/processor.  It does the heavy lifting
 * @returns     Downloaded file to local machine.  Default is Excel file
 * @author      Addam Driver <addam@codedriver.io>
 * @added       01-29-2020
 * @version     1.0
 * @memberOf    kc2excel.js
 * @namespace   grid2Excel
 * @since       1.0
 */
var grid2Excel = (function () {
    let sLocalURI = 'data:application/vnd.ms-excel;base64,'
        , sFinalFileXML = '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">'
            + '<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office"><Author>Axel Richter</Author><Created>{created}</Created></DocumentProperties>'
            + '<Styles>'
            + '<Style ss:ID="Currency"><NumberFormat ss:Format="Currency"></NumberFormat></Style>'
            + '<Style ss:ID="Date"><NumberFormat ss:Format="Medium Date"></NumberFormat></Style>'
            + '</Styles>'
            + '{worksheets}</Workbook>'
        , sWorksheetTemplate = '<Worksheet ss:Name="{nameWS}"><Table>{rows}</Table></Worksheet>'
        , sCellTemplate = '<Cell{attributeStyleID}{attributeFormula}><Data ss:Type="{nameType}">{data}</Data></Cell>'
        /**
         * Method to create the base64 URI component
         */
        , base64 = function (sData) { 
            return window.btoa(unescape(encodeURIComponent(sData))) 
        }
        , fnFormat = function (s, c) { return s.replace(/{(\w+)}/g, function (m, p) { return c[p]; }) }
        /**
         * Method to Create Workbooks (compile) from the HTML Tables
         * @method      fnCreateWorkbooks
         * @param       {array}     aTables             The HTML aTables
         * @param       {array}     aWorkSheetNames     Names of the worksheets
         * @param       {string}    sWorkBookName       Name of the workbook
         * @param       {string}    sApplicationName    Default: Excel - Name of the application (Excel, CSV, etc.)
         * @property    {object}    oConfig             Configuration for sheets and templates
         * @property    {string}    sWorkBookXML        String XML format for a single workbook
         * @property    {string}    sWorkSheetXML       String XML format for a single worksheet
         * @property    {string}    sRowsXML            String XML format for rows and cells
         * @property    {integer}   i                   Iterator for tables
         * @property    {integer}   j                   Iterator for rows
         * @property    {integer}   k                   Iterator for cells
         * @property    {array}     aTablesLen          Cache length of the array tables
         * @property    {array}     aCells              Cache of cell data
         * @property    {number}    nCellLen            Cache of the cell lengths
         * @property    {string}    sDataType           Cell data-type information
         * @property    {string}    sDataStyle          Cell data-style information
         * @property    {string}    sDataValue          Cell data-value information
         * @property    {string}    sDataFormula        Cell data-formula information types
         * @property    {element}   elmLink             HTML link to the local document in the browser
         * @returns     Downloaded file to local machine.  Default is Excel file
         * @author      Addam Driver <addam@codedriver.io>
         * @added       01-29-2020
         * @version     1.0
         * @memberOf    grid2excel
         * @since       1.0
         * @todo        Create a parent div that contains the aTables
                        Loop through the aTables
                        Then remove the div at the very end
         * @example     <caption> Create a workbook from an HTML table </caption>
         * 
         * **** DEFINITION EXAMPLE ****
         * grid2Excel(['TABLEID1', 'TABLEID2'], ['TAB_NAME(FOR TABLEID1)', 'TAB_NAME(FOR TABLEID2)'], 'FILENAME.xls', 'APPLICATION (ie: Excel)')
         * 
         * **** This generates a downloaded file ****
         * // it will look for HTML Table IDs "tbl1" & "tbl2"
         * grid2Excel(['tbl1', 'tbl2'], ['SKUs', 'Locations'], 'TestBook.xls', 'Excel') 
         */
        , fnCreateWorkbooks = function (aTables, aWorkSheetNames, sWorkBookName, sApplicationName = 'Excel') {
            let oConfig = sWorkBookXML = sWorkSheetXML = sRowsXML = ""
                , i = j = k = 0
                , aTableLen = aTables.length || 0
                , aCells
                , nCellLen
                , sDataType
                , sDataStyle
                , sDataValue
                , sDataFormula
                , elmLink
                ;

            // loop through the tables
            for (; i < aTableLen; i += 1) {
                if (!aTables[i].nodeType) {
                    aTables[i] = document.getElementById(aTables[i]);
                }

                // set the counter (needed)
                j = 0;

                // get rows
                for (; j < aTableLen; j += 1) {
                    // set the start of the row XML
                    sRowsXML += '<Row>'

                    // cache the cells
                    console.log(aTables)
                    aCells = aTables[i].rows[j].cells;
                    // cache the cells array length
                    nCellLen = aTables[i].rows[j].cells.length;
                    // set the counter (needed)
                    k = 0;

                    // get aCells
                    for (; k < nCellLen; k++) {
                        //cache the cell values
                        cell = aCells[k];

                        // gather the cell type, style, value & formula
                        sDataType = cell.getAttribute("data-type");
                        sDataStyle = cell.getAttribute("data-style");
                        sDataValue = cell.getAttribute("data-value");
                        sDataFormula = cell.getAttribute("data-formula");

                        // set the value if it exists
                        sDataValue = (sDataValue) ? sDataValue : cell.innerHTML;
                        // set the formula type (calc or date/time)
                        sDataFormula = (sDataFormula) ? sDataFormula : (sApplicationName == 'Calc' && sDataType == 'DateTime') ? sDataValue : null;

                        // set the row oConfig
                        oConfig = {
                            attributeStyleID: (sDataStyle == 'Currency' || sDataStyle == 'Date') ? ' ss:StyleID="' + sDataStyle + '"' : ''
                            , nameType: (sDataType == 'Number' || sDataType == 'DateTime' || sDataType == 'Boolean' || sDataType == 'Error') ? sDataType : 'String'
                            , data: (sDataFormula) ? '' : sDataValue
                            , attributeFormula: (sDataFormula) ? ' ss:Formula="' + sDataFormula + '"' : ''
                        };
                        // complete the cell
                        sRowsXML += fnFormat(sCellTemplate, oConfig);
                    }

                    // complete the row
                    sRowsXML += '</Row>'
                }

                // create the sheet object
                oConfig = {
                    // list of rows
                    rows: sRowsXML,
                    // worksheet name
                    nameWS: aWorkSheetNames[i] || 'Sheet' + i
                };

                // format and include the worksheets
                sWorkSheetXML += fnFormat(sWorksheetTemplate, oConfig);

                // remove the rows and start over
                sRowsXML = "";
            }

            // Meta data for the file
            oConfig = {
                // created date
                created: (new Date()).getTime(),
                // add all of the worksheets
                worksheets: sWorkSheetXML
            };

            // format and create the workbook from the worksheets
            sWorkBookXML = fnFormat(sFinalFileXML, oConfig);

            //console.log(sWorkBookXML);

            // create a link
            elmLink = document.createElement("A");
            // set the internal/local sLocalURI of the workbook
            elmLink.href = sLocalURI + base64(sWorkBookXML);
            // set the name of the workbook/file
            elmLink.download = sWorkBookName || 'Workbook.xls';
            // set the anchor target to nothing
            elmLink.target = '_blank';
            // add the link to the body
            document.body.appendChild(elmLink);
            // simulate a click action on the link
            elmLink.click();
            // remove the link from the DOM
            document.body.removeChild(elmLink);
        }
    ;
    
        
    return fnCreateWorkbooks
})();



var dummy = {
    "SKU_ID": 6597,
    "SKU_NO": "1868-001",
    "SLM_PRJ_ID": 1868,
    "ECC_MATERIAL_NUMBER": "105629900",
    "PLNG_BUS_UNT": "FAM",
    "PRIM_SKU_RL": "B",
    "MDSE_END_DT": null,
    "SKU_SPEC_RQMNTS": null,
    "SKU_LNCH_MABD_DT": "2020-09-24",
    "INIT_SKU_EST_TOTSVOL": null,
    "INIT_SKU_EST_GRS_DLR": null,
    "PROM_PRC_GRP": "PREMIUM-COTTONELLE TP-12DR",
    "RCM_EVDY_RTLR_MRGN": null,
    "FAMA_I_LIC": null,
    "RCM_PROM_RTLR_MRGN": null,
    "ZPU_LST_PRC_UNT": 10.0,
    "TRD_MGMT_PRD_GRP": null,
    "NW_LIC_GRPH_CHAR_I": null,
    "RL_DMNSN_CHG_I": false,
    "SHT_SZ_CHG_I": false,
    "NW_SKU_TYP": "TRDE",
    "MN_VCM_THRSH_SKU_TYP": "BASE",
    "MFG_CMPLX_THRSH_FCTR": "3",
    "MFG_CMPLX_MAX_THRSH": 4.000000,
    "CAP_AVAIL_I": true,
    "LCM_PRD_CHG_I": false,
    "BUOM_SLS_TRND_ASMPTN": "sdf",
    "SU_VOL_MIN_THRSH": 700.000000,
    "VCM_MIN_THRSH": 0.4510,
    "ANUL_OG_XTRACST_CNSD": null,
    "ANUL_OG_XTRACSTS": null,
    "NON_OG_1TM_CST_CNSD": null,
    "NON_OG_1TM_CSTS": null,
    "PKG_ART_PRJ_TYP": "PKGTPB",
    "SKU_RTLR_SMPL_RQ_I": false,
    "PKG_TRIAL_I": false,
    "RLTY_RT_PCT": null,
    "RLVR_INCR_I": "I",
    "TYP_RLVR": null,
    "MEINS": "CS",
    "LABOR": "Z01",
    "GTM_CST_TYP": "FULL",
    "GTM_LMTD_SKU_TYP": [

    ],
    // "CST_RSTRCTNS": [

    // ],
    "ZZNA_FIRST_SHP": "2020-09-10",
    "ZPRDHAL6": "17v1d0",
    "ZPRDHAL6ACDC": "Cottonelle Dry Bath",
    "ZPRDHAL6AGLB": "17v1d0",
    "ZPRDHAL6ARB": "17v1d0",
    "ZPRDHAL1RGN": "200",
    "ZZKC_PRODHIER_1": "200",
    "ZPRDHAL1ABUS": "200",
    "ZZKC_PRODHIER_2": "200",
    "ZPRDHAL2APF": "200",
    "ZPRDHAL2ABR": "200",
    "ZPRDHAL2AES": "200",
    "ZZKC_PRODHIER_3": "1v0",
    "ZPRDHAL3ABG": "1v0",
    "ZPRDHAL3AESS": "1v0",
    "ZZKC_PRODHIER_4": "1i0",
    "ZPRDHAL4ABSG": "1i0",
    "ZPRDHAL5": "17v",
    "ZPRDHAL6ACT": "17v",
    "ZOTCATTR3": false,
    "MATKL": "048",
    "EXTWG": "04800",
    "SUB_BRND_VARNT": "S00024",
    "IRC_I": false,
    "IRC_VAL": null,
    "DVS_ELIG_I": "Y",
    "EXP_20PCT_PKG_CHG_I": false,
    "X_BRND_I": "N",
    "LIC_GRPH_I": false,
    "LIC_GRPH_TYP": "N/A",
    "RL_DIAM": 4.50,
    "SHT_LNTH": 5.000000,
    "SHT_WDTH": 5.000000,
    "FR_GD_SMPL_INCL_I": false,
    "FR_GD_SMPL_TYP": null,
    "FR_GD_IN_V_ON_PK": "N/A",
    "EXP_FR_GD_VLD_END_DT": null,
    "BNDL_I": false,
    "BNDL_CNT": null,
    "FR_GD_SLS_VAL": null,
    "FR_GD_SLS_VALPCT": null,
    "ZRTLPKCNT": 10,
    "RTL_CS_I": false,
    "MDSE_CONFIG": "MC1",
    "MDSE_FTPRNT": "MF5",
    "CMPNT_I": false,
    "ZOTCATTR5": null,
    "MTART": null,
    "ZEAPCKIND": "N",
    "ZGRP1PLCT": null,
    "ZGRP2PLCT": null,
    "ZGRP3PLCT": null,
    "SLS_BOM_CMPNT_TOT_PCS": 0.000,
    "PLNG_VRNT": null,
    // "SRC_OF_SPLY_TYP": [
    //     {
    //         "value": "SRC3"
    //     },
    //     {
    //         "value": "SRC3"
    //     }
    // ],
    "ZGSUFCTOR": 10000.000000,
    "ZGSUOVRD": null,
    "GSU_PR_RTL_PKG": 0.020000,
    "GSU_PR_BUOM": 0.200000,
    "LST_PRC_RPKG_CAD": null,
    "CPY_SKU_VERS_PRD_CD": "104752300",
    "ZPRDHAL7": "17v1d0ll",
    "ZPRDHAL8": "17v1d0ll192j4",
    "ZPRDHAL9": "17v1d0ll192j430",
    "ZPRDHAL10": "17v1d0ll192j430600",
    "SPART": "00",
    "FZPLP3": 5.49,
    "ADD_ON_PROM_I": false,
    "SPCL_PK_DIS_MENU_I": false,
    "SIOC_CMPAT_I": "N/A",
    // "CPK_TYP": [
    //     {
    //         "value": "N/A"
    //     },
    //     {
    //         "value": "N/A"
    //     }
    // ],
    "EST_RPKG_HGHT": 15.000,
    "EST_RPKG_LEN": 20.000,
    "EST_RPKG_WID": 22.000,
    "FLX_GRPH_I": true,
    "CS_PK_OUT_CMPLY_I": "Y",
    "RPKG_GRPH_LANG": "LNG2",
    "ZOTCCHAR4": "Z0059",
    "ZRTPKMTYP": "PL",
    "ZLICHCDNM": null,
    "LIC_MDSE_CAT_CD": null,
    "VRG": null,
    "STD_V_NL_ALLPCT": null,
    "LIC_CHAR_V_NL_ALLPCT": null,
    "FAMA_I": null,
    "LIC_GRPH_EXP_APRV_I": "N/A",
    "LIC_GRPH_IMP_APRV_I": null,
    "EXP_ONLY_SLS_I": "N",
    // "ZCTRYEXSL": [

    // ],
    // "HERKL": [
    //     {
    //         "value": "US"
    //     }
    // ],
    "ZECCNNUM": "EAR99",
    "PROM_MTL_EXP_APRV_I": null,
    "FSC_CERTIFIED_INDICATOR": null,
    "FSC_CRT_NBR": null,
    "XCHPFMARA": "N",
    "KTGRM": "M1",
    "MVGR5": "TLD",
    "BASE_BAF_RT_CANADA": null,
    "ZDDMGRPID": "61",
    "ZEWMINDIC": false,
    "ZCPNFMYCD": "670",
    "CPN_REDMP_RT_PCT": null,
    "EST_OG_FX_COM_DLR": 1.21,
    "EST_OG_VR_COM_DLR": 4.28,
    "CNSMR_SMPL_VAL": null,
    "SKU_BSLN_COM_DLR_CHG": null,
    "TXTPURCH": "COTT B+R BT 10 PK 200 ",
    "MFRNR": null,
    "MFRPN": null,
    "EKGRP": null,
    "KAUTB": null,
    "ZTRDITEM": "991056299",
    "MATNR": "105629900",
    "CYCL_TGT_I": null,
    "INTRNS_TGT_I": null,
    "ZAPOCHR19": "S23",
    "GNDR": "N/A",
    "AST_TECH1": null,
    "APO_MTL_SEL1": null,
    "APO_MTL_SEL2": null,
    "ABSRBNCY": "N/A",
    "ZSKUSUMMR": "Cottonelle B+R BATHROOM TISSUE  10 PK 200\r\nEN/FR; Case; MTS\r\nU.S.\r\nGSU= 0.2; 10Rtl Pkg per CS; 200Prd Cnt per Rtl Pkg\r\nCTNLDR Tracy 1-13\r\nMade in  US",
    "ZSKUSUMMR_AUTO": null,
    "ZOTCCHAR5": "MTS",
    "CLASSTYPE": null,
    "ZGDSRELEV": true,
    "ZEQVLNTCS": "1",
    "TXTMI": "COTT B+R BT 10 PK 200 ",
    "SKU_LONG_DESC": "Cottonelle B+R BATHROOM TISSUE  10 PK 200",
    "NOTEBSCDA": "Cottonelle B+R BATHROOM TISSUE  10 PK 200\r\nEN/FR; Case; MTS\r\nU.S.",
    "TXTSALES": "Cottonelle BIG PLUS ROLL BATHROOM TISSUE 10 PACK 200 ",
    "TXTTAPE": "COTT BT 10 PK 200",
    "ZBRNDNM": "COTT",
    "ZSTYLNM": "B+R",
    "ZCNTNTNM": "BT",
    "ZCOLOR": "-",
    "ZSZNM": "10 PK",
    "ZCDMCDE": "03",
    "DP_PACKING": null,
    "ZIBPDP": null,
    "ZIBPDPIND": "False",
    "WTCHDG_I": false,
    "REGION": null,
    "PROJECT_SUB_STATUS": null,
    "CREATED_TIMESTAMP": "2020-01-13T15:17:38.99",
    "CREATED_USERID": "B20896",
    "LAST_UPD_TIMESTMP": "2020-01-27T17:33:15.92",
    "LAST_UPD_USERID": "B20896",
    "CAS_ADJ_FACTOR": "1.000000",
    "CAS_I": false,
    "CMNT_CODE_GSUO": null,
    "ECOM_SORT_CMPLY_I": "Y",
    "FR_GD_BNDL_CNT": null,
    "FR_GD_PRD_CNT": null,
    "GSU_CALC_PR_BUOM": 0.200000,
    "LIKE_TRD_SKU": "991047523",
    "NEW_STYLE_EPH_I": false,
    "PRD_CNT": 200,
    "PRE_PRC_I": false,
    "PRE_PRC_VAL": null,
    "SALES_BOM_CMPNT_CMNTS": null,
    "SKU_COUNTER": "1868-001",
    "ZCHFSCCRT": "N",
    "ZEPARGIND": false,
    "ZOVRCTIND": false,
    "ZPRDTCNT": "200",
    "ZSTKHGMET": null,
    "CATMAN_PKD_ENDRS_I": "Y",
    "MTL_RESERV_I": "Y",
    "RSRV_MTL_CMPL": true,
    "CPY_TRDE_FIXCST_BSLN": 1.21,
    "CPY_TRDE_TOTCOM_BSLN": 5.49,
    "CPY_TRDE_VARCST_BSLN": 4.28,
    "SKU_SLM_PLAT": "CTNLDR",
    "GTIN_EXCPTN_APRV_I": "N/A",
    "LIC_GRPH_DOM_APRV_I": null,
    "LIC_GRPH_EXP_APRV_I_LIC": "N/A",
    "LIC_GRPH_IMP_APRV_I_LIC": null,
    "LIC_GRPH_DOM_APRV_I_LIC": null,
    "SKU_LVL_HLTH_PLAT": "CTNLDR",
    "SKU_ULD_AVAIL_I": true,
    // "SalesBOM_DATA": [

    // ],
    // "COUNTRY_DATA": [
    //     {
    //         "SKU_CON_ID": 8459,
    //         "SKU_ID": 6597,
    //         "CTY_CODE": "US",
    //         "AVG_BAF_RT": 0.1940,
    //         "PERF_TIER_BAF_RT": 0.0240,
    //         "STRAT_TIER_BAF_RT": 0.0400,
    //         "WGHT_AVG_BAF_RT": 0.0500,
    //         "MSRP_EVDY": 8.99,
    //         "RCM_EVDY_NETRTLR_PRC": null,
    //         "MSRP_PROM": 6.49,
    //         "RCM_PROM_NETRTLR_PRC": null,
    //         "EX_P_V_NP_PCTVOL_SPL": null,
    //         "MSRP_AVG": null,
    //         "NET_RTLR_PRC": null,
    //         "LAC": 5.16,
    //         "LST_PRC_EFTV_DT": "2020-06-22",
    //         "LST_PRC_EFTV_END_DT": "9999-12-31",
    //         "LST_PRC_PR_RPKG_USD": 7.69,
    //         "ZPU_LST_PRC": 7.69,
    //         "BUOM_LST_PRC": 76.90,
    //         "LST_PRC_PCT_OV": null,
    //         "LST_PRC_PCT_OV_CMTS": null,
    //         "CALC_BASE_LIST_PRICE": 76.90,
    //         "FR_GD_SLS_VAL": null,
    //         "FR_GD_SLS_VALPCT": 0.000,
    //         "CPY_TRD_ITM_ACT_SVOL": 0.0,
    //         "NBR_WEEK_W_SLS": 0,
    //         "BUOM_SLS_TRND": 0.0000,
    //         "SKU_TOT_SVOL": 0.0,
    //         "SKU_TOT_GRS_DLR": 0.00,
    //         "PCT_INCR_VOL": 0.3000,
    //         "SKU_TOT_SVOL_INCR": 0.0,
    //         "SKU_GRS_DLR_INCR": 0.00,
    //         "CPY_TRDE_VCMPCT_BSLN": 0.0000,
    //         "SKU_ABS_VCM_PCT_CHG": null,
    //         "VCMPCT_EST_H": 0.1000,
    //         "VCMPCT_EST_L": 0.0500,
    //         "VCMPCT_EST_AVG": 0.0750,
    //         "MVGR2": null,
    //         "MVGR3": null,
    //         "STAMARCFT": null,
    //         "CREATED_TIMESTAMP": "2020-01-13T15:17:42.083",
    //         "CREATED_USERID": "B20896",
    //         "LAST_UPD_TIMESTMP": "2020-01-23T18:22:03.707",
    //         "LAST_UPD_USERID": "B20896"
    //     }
    // ],
    // "LOCATION_DATA": [
    //     {
    //         "SKU_LOC_ID": 22431,
    //         "SKU_ID": 6597,
    //         "WERKS": "2019",
    //         "WZEIT": null,
    //         "BESKZ": "F",
    //         "ZMNFTRPLT": "N",
    //         "ZFRCTHORZ": null,
    //         "VNDR_LT": null,
    //         "BIG_RCVG_LOC": "N",
    //         "ZGLPROTIM": 1.0,
    //         "ZUNLTOVDL": "False",
    //         "LOCATION_DATA": null,
    //         "FAMA_I": null,
    //         "CREATED_TIMESTAMP": "2020-01-22T18:18:19.133",
    //         "CREATED_USERID": "B20896",
    //         "LAST_UPD_TIMESTMP": "2020-01-23T17:14:03.057",
    //         "LAST_UPD_USERID": "B20896"
    //     },
    //     {
    //         "SKU_LOC_ID": 22465,
    //         "SKU_ID": 6597,
    //         "WERKS": "2027",
    //         "WZEIT": null,
    //         "BESKZ": "X",
    //         "ZMNFTRPLT": "Y",
    //         "ZFRCTHORZ": null,
    //         "VNDR_LT": null,
    //         "BIG_RCVG_LOC": "N",
    //         "ZGLPROTIM": 1.0,
    //         "ZUNLTOVDL": "False",
    //         "LOCATION_DATA": null,
    //         "FAMA_I": null,
    //         "CREATED_TIMESTAMP": "2020-01-22T18:18:19.133",
    //         "CREATED_USERID": "B20896",
    //         "LAST_UPD_TIMESTMP": "2020-01-23T17:13:54.6",
    //         "LAST_UPD_USERID": "B20896"
    //     },
    //     {
    //         "SKU_LOC_ID": 22432,
    //         "SKU_ID": 6597,
    //         "WERKS": "2031",
    //         "WZEIT": null,
    //         "BESKZ": "F",
    //         "ZMNFTRPLT": "N",
    //         "ZFRCTHORZ": null,
    //         "VNDR_LT": null,
    //         "BIG_RCVG_LOC": "N",
    //         "ZGLPROTIM": 1.0,
    //         "ZUNLTOVDL": "False",
    //         "LOCATION_DATA": null,
    //         "FAMA_I": null,
    //         "CREATED_TIMESTAMP": "2020-01-22T18:18:19.133",
    //         "CREATED_USERID": "B20896",
    //         "LAST_UPD_TIMESTMP": "2020-01-23T17:14:02.29",
    //         "LAST_UPD_USERID": "B20896"
    //     },
    //     {
    //         "SKU_LOC_ID": 22433,
    //         "SKU_ID": 6597,
    //         "WERKS": "2050",
    //         "WZEIT": null,
    //         "BESKZ": "F",
    //         "ZMNFTRPLT": "N",
    //         "ZFRCTHORZ": null,
    //         "VNDR_LT": null,
    //         "BIG_RCVG_LOC": "N",
    //         "ZGLPROTIM": 1.0,
    //         "ZUNLTOVDL": "False",
    //         "LOCATION_DATA": null,
    //         "FAMA_I": null,
    //         "CREATED_TIMESTAMP": "2020-01-22T18:18:19.133",
    //         "CREATED_USERID": "B20896",
    //         "LAST_UPD_TIMESTMP": "2020-01-23T17:14:01.523",
    //         "LAST_UPD_USERID": "B20896"
    //     },
    //     {
    //         "SKU_LOC_ID": 22435,
    //         "SKU_ID": 6597,
    //         "WERKS": "2100",
    //         "WZEIT": null,
    //         "BESKZ": "F",
    //         "ZMNFTRPLT": "N",
    //         "ZFRCTHORZ": null,
    //         "VNDR_LT": null,
    //         "BIG_RCVG_LOC": "N",
    //         "ZGLPROTIM": 1.0,
    //         "ZUNLTOVDL": "False",
    //         "LOCATION_DATA": null,
    //         "FAMA_I": null,
    //         "CREATED_TIMESTAMP": "2020-01-22T18:18:19.133",
    //         "CREATED_USERID": "B20896",
    //         "LAST_UPD_TIMESTMP": "2020-01-23T17:14:00.007",
    //         "LAST_UPD_USERID": "B20896"
    //     },
    //     {
    //         "SKU_LOC_ID": 22430,
    //         "SKU_ID": 6597,
    //         "WERKS": "2169",
    //         "WZEIT": null,
    //         "BESKZ": "F",
    //         "ZMNFTRPLT": "N",
    //         "ZFRCTHORZ": null,
    //         "VNDR_LT": null,
    //         "BIG_RCVG_LOC": "N",
    //         "ZGLPROTIM": 1.0,
    //         "ZUNLTOVDL": "False",
    //         "LOCATION_DATA": null,
    //         "FAMA_I": null,
    //         "CREATED_TIMESTAMP": "2020-01-22T18:18:19.133",
    //         "CREATED_USERID": "B20896",
    //         "LAST_UPD_TIMESTMP": "2020-01-23T17:13:53.053",
    //         "LAST_UPD_USERID": "B20896"
    //     },
    //     {
    //         "SKU_LOC_ID": 22436,
    //         "SKU_ID": 6597,
    //         "WERKS": "2292",
    //         "WZEIT": null,
    //         "BESKZ": "F",
    //         "ZMNFTRPLT": "N",
    //         "ZFRCTHORZ": null,
    //         "VNDR_LT": null,
    //         "BIG_RCVG_LOC": "N",
    //         "ZGLPROTIM": 1.0,
    //         "ZUNLTOVDL": "False",
    //         "LOCATION_DATA": null,
    //         "FAMA_I": null,
    //         "CREATED_TIMESTAMP": "2020-01-22T18:18:19.133",
    //         "CREATED_USERID": "B20896",
    //         "LAST_UPD_TIMESTMP": "2020-01-23T17:13:59.24",
    //         "LAST_UPD_USERID": "B20896"
    //     },
    //     {
    //         "SKU_LOC_ID": 22437,
    //         "SKU_ID": 6597,
    //         "WERKS": "2299",
    //         "WZEIT": null,
    //         "BESKZ": "F",
    //         "ZMNFTRPLT": "N",
    //         "ZFRCTHORZ": null,
    //         "VNDR_LT": null,
    //         "BIG_RCVG_LOC": "N",
    //         "ZGLPROTIM": 1.0,
    //         "ZUNLTOVDL": "False",
    //         "LOCATION_DATA": null,
    //         "FAMA_I": null,
    //         "CREATED_TIMESTAMP": "2020-01-22T18:18:19.133",
    //         "CREATED_USERID": "B20896",
    //         "LAST_UPD_TIMESTMP": "2020-01-23T17:13:58.49",
    //         "LAST_UPD_USERID": "B20896"
    //     },
    //     {
    //         "SKU_LOC_ID": 22438,
    //         "SKU_ID": 6597,
    //         "WERKS": "2320",
    //         "WZEIT": null,
    //         "BESKZ": "F",
    //         "ZMNFTRPLT": "N",
    //         "ZFRCTHORZ": null,
    //         "VNDR_LT": null,
    //         "BIG_RCVG_LOC": "N",
    //         "ZGLPROTIM": 1.0,
    //         "ZUNLTOVDL": "False",
    //         "LOCATION_DATA": null,
    //         "FAMA_I": null,
    //         "CREATED_TIMESTAMP": "2020-01-22T18:18:19.133",
    //         "CREATED_USERID": "B20896",
    //         "LAST_UPD_TIMESTMP": "2020-01-23T17:13:57.71",
    //         "LAST_UPD_USERID": "B20896"
    //     },
    //     {
    //         "SKU_LOC_ID": 22439,
    //         "SKU_ID": 6597,
    //         "WERKS": "2336",
    //         "WZEIT": null,
    //         "BESKZ": "F",
    //         "ZMNFTRPLT": "N",
    //         "ZFRCTHORZ": null,
    //         "VNDR_LT": null,
    //         "BIG_RCVG_LOC": "N",
    //         "ZGLPROTIM": 1.0,
    //         "ZUNLTOVDL": "False",
    //         "LOCATION_DATA": null,
    //         "FAMA_I": null,
    //         "CREATED_TIMESTAMP": "2020-01-22T18:18:19.133",
    //         "CREATED_USERID": "B20896",
    //         "LAST_UPD_TIMESTMP": "2020-01-23T17:13:56.927",
    //         "LAST_UPD_USERID": "B20896"
    //     },
    //     {
    //         "SKU_LOC_ID": 22425,
    //         "SKU_ID": 6597,
    //         "WERKS": "2358",
    //         "WZEIT": null,
    //         "BESKZ": "F",
    //         "ZMNFTRPLT": "N",
    //         "ZFRCTHORZ": null,
    //         "VNDR_LT": null,
    //         "BIG_RCVG_LOC": "N",
    //         "ZGLPROTIM": 1.0,
    //         "ZUNLTOVDL": "False",
    //         "LOCATION_DATA": null,
    //         "FAMA_I": null,
    //         "CREATED_TIMESTAMP": "2020-01-22T18:18:19.133",
    //         "CREATED_USERID": "B20896",
    //         "LAST_UPD_TIMESTMP": "2020-01-23T17:13:52.253",
    //         "LAST_UPD_USERID": "B20896"
    //     },
    //     {
    //         "SKU_LOC_ID": 22434,
    //         "SKU_ID": 6597,
    //         "WERKS": "2360",
    //         "WZEIT": null,
    //         "BESKZ": "F",
    //         "ZMNFTRPLT": "N",
    //         "ZFRCTHORZ": null,
    //         "VNDR_LT": null,
    //         "BIG_RCVG_LOC": "N",
    //         "ZGLPROTIM": 1.0,
    //         "ZUNLTOVDL": "False",
    //         "LOCATION_DATA": null,
    //         "FAMA_I": null,
    //         "CREATED_TIMESTAMP": "2020-01-22T18:18:19.133",
    //         "CREATED_USERID": "B20896",
    //         "LAST_UPD_TIMESTMP": "2020-01-23T17:14:00.773",
    //         "LAST_UPD_USERID": "B20896"
    //     },
    //     {
    //         "SKU_LOC_ID": 22440,
    //         "SKU_ID": 6597,
    //         "WERKS": "2369",
    //         "WZEIT": null,
    //         "BESKZ": "F",
    //         "ZMNFTRPLT": "N",
    //         "ZFRCTHORZ": null,
    //         "VNDR_LT": null,
    //         "BIG_RCVG_LOC": "N",
    //         "ZGLPROTIM": 1.0,
    //         "ZUNLTOVDL": "False",
    //         "LOCATION_DATA": null,
    //         "FAMA_I": null,
    //         "CREATED_TIMESTAMP": "2020-01-22T18:18:19.133",
    //         "CREATED_USERID": "B20896",
    //         "LAST_UPD_TIMESTMP": "2020-01-23T17:13:56.18",
    //         "LAST_UPD_USERID": "B20896"
    //     },
    //     {
    //         "SKU_LOC_ID": 22457,
    //         "SKU_ID": 6597,
    //         "WERKS": "2422",
    //         "WZEIT": null,
    //         "BESKZ": "F",
    //         "ZMNFTRPLT": "N",
    //         "ZFRCTHORZ": null,
    //         "VNDR_LT": null,
    //         "BIG_RCVG_LOC": "N",
    //         "ZGLPROTIM": 1.0,
    //         "ZUNLTOVDL": "False",
    //         "LOCATION_DATA": null,
    //         "FAMA_I": null,
    //         "CREATED_TIMESTAMP": "2020-01-22T18:18:19.133",
    //         "CREATED_USERID": "B20896",
    //         "LAST_UPD_TIMESTMP": "2020-01-23T17:13:51.457",
    //         "LAST_UPD_USERID": "B20896"
    //     },
    //     {
    //         "SKU_LOC_ID": 22441,
    //         "SKU_ID": 6597,
    //         "WERKS": "2496",
    //         "WZEIT": null,
    //         "BESKZ": "F",
    //         "ZMNFTRPLT": "N",
    //         "ZFRCTHORZ": null,
    //         "VNDR_LT": null,
    //         "BIG_RCVG_LOC": "N",
    //         "ZGLPROTIM": 1.0,
    //         "ZUNLTOVDL": "False",
    //         "LOCATION_DATA": null,
    //         "FAMA_I": null,
    //         "CREATED_TIMESTAMP": "2020-01-22T18:18:19.133",
    //         "CREATED_USERID": "B20896",
    //         "LAST_UPD_TIMESTMP": "2020-01-23T17:13:55.38",
    //         "LAST_UPD_USERID": "B20896"
    //     },
    //     {
    //         "SKU_LOC_ID": 22443,
    //         "SKU_ID": 6597,
    //         "WERKS": "2833",
    //         "WZEIT": null,
    //         "BESKZ": "X",
    //         "ZMNFTRPLT": "Y",
    //         "ZFRCTHORZ": null,
    //         "VNDR_LT": null,
    //         "BIG_RCVG_LOC": "N",
    //         "ZGLPROTIM": 1.0,
    //         "ZUNLTOVDL": "False",
    //         "LOCATION_DATA": null,
    //         "FAMA_I": null,
    //         "CREATED_TIMESTAMP": "2020-01-22T18:18:19.133",
    //         "CREATED_USERID": "B20896",
    //         "LAST_UPD_TIMESTMP": "2020-01-23T17:13:53.817",
    //         "LAST_UPD_USERID": "B20896"
    //     }
    // ],
    // "SALESORG_DATA": [
    //     {
    //         "SKU_SAL_ORG_ID": 8563,
    //         "SKU_ID": 6597,
    //         "VKORG": "2810",
    //         "VTWEG": "80",
    //         "ZBUKRS": "0008",
    //         "SKU_GL_SLS_RGN": "CNSUS",
    //         "COUNTRY": "US",
    //         "CREATED_TIMESTAMP": "2020-01-13T15:17:42.083",
    //         "CREATED_USERID": "B20896",
    //         "LAST_UPD_TIMESTMP": "2020-01-13T15:17:42.083",
    //         "LAST_UPD_USERID": "B20896"
    //     }
    // ],
    // "SELECTION_DATA_DESC": [
    //     {
    //         "fieldKey": "GNDR",
    //         "selection": "N/A",
    //         "selectionDesc": "N/A"
    //     },
    //     {
    //         "fieldKey": "GTM_CST_TYP",
    //         "selection": "FULL",
    //         "selectionDesc": "Full National Availability"
    //     },
    //     {
    //         "fieldKey": "LIC_GRPH_TYP",
    //         "selection": "N/A",
    //         "selectionDesc": "N/A"
    //     },
    //     {
    //         "fieldKey": "MDSE_CONFIG",
    //         "selection": "MC1",
    //         "selectionDesc": "Case"
    //     },
    //     {
    //         "fieldKey": "MDSE_FTPRNT",
    //         "selection": "MF5",
    //         "selectionDesc": "N/A"
    //     },
    //     {
    //         "fieldKey": "NW_SKU_TYP",
    //         "selection": "TRDE",
    //         "selectionDesc": "New Trade Item"
    //     },
    //     {
    //         "fieldKey": "PKG_ART_PRJ_TYP",
    //         "selection": "PKGTPB",
    //         "selectionDesc": "New design"
    //     },
    //     {
    //         "fieldKey": "PRIM_SKU_RL",
    //         "selection": "B",
    //         "selectionDesc": "Base (EveryDay) "
    //     },
    //     {
    //         "fieldKey": "RLVR_INCR_I",
    //         "selection": "I",
    //         "selectionDesc": "Incremental"
    //     },
    //     {
    //         "fieldKey": "RPKG_GRPH_LANG",
    //         "selection": "LNG2",
    //         "selectionDesc": "English/French"
    //     },
    //     {
    //         "fieldKey": "SLM_PLAT",
    //         "selection": "CTNLDR",
    //         "selectionDesc": "Cottonelle Dry"
    //     },
    //     {
    //         "fieldKey": "SUB_BRND_VARNT",
    //         "selection": "S00024",
    //         "selectionDesc": "Ultra Clean Care"
    //     },
    //     {
    //         "fieldKey": "DVS_ELIG_I",
    //         "selection": "Y",
    //         "selectionDesc": "Yes - DVS Eligible"
    //     },
    //     {
    //         "fieldKey": "CS_PK_OUT_CMPLY_I",
    //         "selection": "Y",
    //         "selectionDesc": "Yes - Checked"
    //     },
    //     {
    //         "fieldKey": "SIOC_CMPAT_I",
    //         "selection": "N/A",
    //         "selectionDesc": "Not Applicable"
    //     },
    //     {
    //         "fieldKey": "CATMAN_PKD_ENDRS_I",
    //         "selection": "Y",
    //         "selectionDesc": "Yes - Endorsed"
    //     },
    //     {
    //         "fieldKey": "LIC_GRPH_EXP_APRV_I",
    //         "selection": "N/A",
    //         "selectionDesc": "Not Applicable"
    //     },
    //     {
    //         "fieldKey": "FR_GD_IN_V_ON_PK",
    //         "selection": "N/A",
    //         "selectionDesc": "Not Applicable"
    //     },
    //     {
    //         "fieldKey": "ECOM_SORT_CMPLY_I",
    //         "selection": "Y",
    //         "selectionDesc": "Yes - compliant"
    //     },
    //     {
    //         "fieldKey": "ZEAPCKIND",
    //         "selection": "N",
    //         "selectionDesc": "FALSE"
    //     },
    //     {
    //         "fieldKey": "XCHPFMARA",
    //         "selection": "N",
    //         "selectionDesc": "No"
    //     },
    //     {
    //         "fieldKey": "ZCPNFMYCD",
    //         "selection": "670",
    //         "selectionDesc": "Cottonelle"
    //     },
    //     {
    //         "fieldKey": "MVGR5",
    //         "selection": "TLD",
    //         "selectionDesc": "Truckload"
    //     },
    //     {
    //         "fieldKey": "ZOTCCHAR5",
    //         "selection": "MTS",
    //         "selectionDesc": "Make to Stock"
    //     },
    //     {
    //         "fieldKey": "GTIN_EXCPTN_APRV_I",
    //         "selection": "N/A",
    //         "selectionDesc": "Not Applicable"
    //     },
    //     {
    //         "fieldKey": "ZCNTNTNM",
    //         "selection": "BT",
    //         "selectionDesc": "BATHROOM TISSUE"
    //     },
    //     {
    //         "fieldKey": "ZSZNM",
    //         "selection": "10 PK",
    //         "selectionDesc": "10 PACK"
    //     },
    //     {
    //         "fieldKey": "ZSTYLNM",
    //         "selection": "B+R",
    //         "selectionDesc": "BIG PLUS ROLL"
    //     },
    //     {
    //         "fieldKey": "ZBRNDNM",
    //         "selection": "COTT",
    //         "selectionDesc": "Cottonelle"
    //     },
    //     {
    //         "fieldKey": "ZCOLOR",
    //         "selection": "-",
    //         "selectionDesc": "COLOR NOT APPLICABLE"
    //     },
    //     {
    //         "fieldKey": "ABSRBNCY",
    //         "selection": "N/A",
    //         "selectionDesc": "N/A"
    //     },
    //     {
    //         "fieldKey": "MTL_RESERV_I",
    //         "selection": "Y",
    //         "selectionDesc": "Yes"
    //     },
    //     {
    //         "fieldKey": "MFG_CMPLX_THRSH_FCTR",
    //         "selection": "3",
    //         "selectionDesc": "Maximum Threshold 3"
    //     },
    //     {
    //         "fieldKey": "SKU_SLM_PLAT",
    //         "selection": "CTNLDR",
    //         "selectionDesc": "Cottonelle Dry"
    //     },
    //     {
    //         "fieldKey": "PROM_PRC_GRP",
    //         "selection": "PREMIUM-COTTONELLE TP-12DR",
    //         "selectionDesc": "PREMIUM-COTTONELLE TP-12DR"
    //     }
    // ],
    // "EPH_SELECTION_DATA_DESC": [
    //     {
    //         "fieldKey": "ZZKC_PRODHIER_1",
    //         "selection": "200",
    //         "selectionDesc": "Family Care"
    //     },
    //     {
    //         "fieldKey": "ZPRDHAL1RGN",
    //         "selection": "GL",
    //         "selectionDesc": "GL"
    //     },
    //     {
    //         "fieldKey": "ZPRDHAL1ABUS",
    //         "selection": "Consumer",
    //         "selectionDesc": "Consumer"
    //     },
    //     {
    //         "fieldKey": "ZZKC_PRODHIER_2",
    //         "selection": "200",
    //         "selectionDesc": "Branded Perineal Hygiene"
    //     },
    //     {
    //         "fieldKey": "ZPRDHAL2AES",
    //         "selection": "Family Care",
    //         "selectionDesc": "Family Care"
    //     },
    //     {
    //         "fieldKey": "ZPRDHAL2ABR",
    //         "selection": "Y",
    //         "selectionDesc": "Y"
    //     },
    //     {
    //         "fieldKey": "ZPRDHAL2APF",
    //         "selection": "Perineal Hygiene",
    //         "selectionDesc": "Perineal Hygiene"
    //     },
    //     {
    //         "fieldKey": "ZZKC_PRODHIER_3",
    //         "selection": "1v0",
    //         "selectionDesc": "Branded Perineal Hygiene"
    //     },
    //     {
    //         "fieldKey": "ZPRDHAL3ABG",
    //         "selection": "Perineal Hygiene",
    //         "selectionDesc": "Perineal Hygiene"
    //     },
    //     {
    //         "fieldKey": "ZPRDHAL3AESS",
    //         "selection": "Family Care",
    //         "selectionDesc": "Family Care"
    //     },
    //     {
    //         "fieldKey": "ZZKC_PRODHIER_4",
    //         "selection": "1i0",
    //         "selectionDesc": "Branded Dry Bath"
    //     },
    //     {
    //         "fieldKey": "ZPRDHAL4ABSG",
    //         "selection": "Dry Bath",
    //         "selectionDesc": "Dry Bath"
    //     },
    //     {
    //         "fieldKey": "ZPRDHAL5",
    //         "selection": "17v",
    //         "selectionDesc": "Branded Dry Bath"
    //     },
    //     {
    //         "fieldKey": "ZPRDHAL6ACT",
    //         "selection": "Dry Bath",
    //         "selectionDesc": "Dry Bath"
    //     },
    //     {
    //         "fieldKey": "ZPRDHAL6",
    //         "selection": "17v1d0",
    //         "selectionDesc": "Cottonelle Dry Bath"
    //     },
    //     {
    //         "fieldKey": "ZPRDHAL6ACDC",
    //         "selection": "Cottonelle Dry Bath",
    //         "selectionDesc": "Cottonelle Dry Bath"
    //     },
    //     {
    //         "fieldKey": "ZPRDHAL6ARB",
    //         "selection": "Cottonelle",
    //         "selectionDesc": "Cottonelle"
    //     },
    //     {
    //         "fieldKey": "ZPRDHAL6AGLB",
    //         "selection": "Cottonelle",
    //         "selectionDesc": "Cottonelle"
    //     },
    //     {
    //         "fieldKey": "ZPRDHAL7",
    //         "selection": "17v1d0ll",
    //         "selectionDesc": "Cottonelle Aloe & E/GentleCare Dry Bath"
    //     },
    //     {
    //         "fieldKey": "ZPRDHAL8",
    //         "selection": "17v1d0ll192j4",
    //         "selectionDesc": "Cottonelle Aloe & E Dry Bath Big Roll"
    //     },
    //     {
    //         "fieldKey": "ZPRDHAL9",
    //         "selection": "17v1d0ll192j430",
    //         "selectionDesc": "Cottonelle Aloe & E Dry Bath BR P8"
    //     },
    //     {
    //         "fieldKey": "ZPRDHAL10",
    //         "selection": "17v1d0ll192j430600",
    //         "selectionDesc": "COTT A&E BR 8PK255-260"
    //     }
    // ],
    // "LICENSEDGRAPHIC_DATA": [

    // ],
    // "STUDIOLEVEL_DATA": [

    // ]
}

/**
 * WIP:  Method to create an HTML table on the page.
 * It will take a flattened object (2 levels) and if any level is an array of data,
 * It should create a "TAB" and an additional table.  Think recursion
 * @method  createTable
 */
let createTable = function (aData) {
    let elem = document.getElementById('myTable')
    , final = ''
    //, tableStart = '<table border=1 id="tbl0">'
    , tableend = '</table>'
    , headers = []
    , wbData = ['tbl0']
    , wbNames = ['main']
    , i = 0
    , setRows = function(aData) {
        let final = ''
        //, headers = []
        , tr = ''
        , td = ''
        , th = ''
        ;

        for(let i = 0; i < aData.length; i += 1) {
                
            for (let key in aData[i]) {
                
                // check for headers
                if (headers.indexOf(key) === -1) {
                    headers.push(key);
                    tr += '<th>' + key + '</th>';
                    // console.log(key)
                } 
                //console.warn(data[i][key])
                // else {
                //     // get cell data
                //     // td += '<td>' + data[i][key] + '</td>'
                // }
            }
            
             
            
        }
        //console.log(headers)
        return tr; 
    }
    ;
    let shTr = '<tr>'
    , sTh = '<tr>'
    , tableStart = '<table border=1 id="tbl'+i+'">'
    ;


    for (var key in dummy) {
        if (Array.isArray(dummy[key])) {
            // store the first part of data
             
                 wbData.push("tbl" + i);
                 // add the tab name
                 wbNames.push(key);
                 i += 1;
             
            
            //final += setRows(dummy[key]);
            break;
        } else {
            if (headers.indexOf(key) === -1) {
                headers.push(key);
                shTr += '<th>' + key + '</th>';
                // console.log(key)
            } 
            sTh += '<td>' + dummy[key] + '</td>';
        }
        
        
        
        // if (headers.indexOf(key) == -1) {
        //     headers.push(key);
        //     final += '<th>' + key + '</th>';
            
        // }
    }
    shTr += '</tr>';
    sTh += '</tr>';
    final += shTr + sTh;
    console.log(wbData, wbNames, headers);
    
    elem.innerHTML = tableStart + final + tableend;
    
    
    shTr = ''
    sTh = ''
    final=''
    
    grid2Excel(wbData, wbNames, 'testBook.xls')
    
}
console.log(dummy)
createTable(dummy);



//grid2Excel(['tbl1', 'tbl2', 'tbl3'], ['Customers', 'Products', 'junk'], 'TestBook.xls', 'Excel')