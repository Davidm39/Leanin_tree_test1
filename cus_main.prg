
#INCLUDE "c_Cust.Def"
****************************************************************************
#define ANSI_CR_LF  chr(13)+chr(10)
****************************************************************************

*CUS_MAIN is used to store used defined functions.
*This program file will be loaded during system initialization so that
*all functions within the file will be available for the user to access
*via PowerAct execute VFP functions.
*The function CUS_INIT will be called during system initialization
*The function CUS_END  will be called when user exits the system
*
*The user can add his own functions and also SET PROCEDURE to another
*FXP files that include more functions.
****************************************************************************
*!* I went thru every line, should be good to go
****************************************************************************
FUNCTION Cus_Init()
****************************************************************************
*wait window "Started with new Cus_main-code for future shipments" timeout.25
* Add here code to be executed after Login Screen during Ascent Entry
* Return .F. to Cancel Ascent Execution
wait window "MSEC CHANGES-FOR SAP UPGRADE 01/01/09 " timeout 1 
copyshipper()
** 3/1/2004  changed time format for output 
** 06/12/03  CHANGED VOID FUNCTION FOR TRADE

Public gcORDER, gcShipVia, gcTrknum, gcsycd ,glFut, gnMultCOD 
PUBLIC gcGander, gcService
gcGander = ""
gcTrknum = ""	
gcoRDER = ""	&& VBELN
gcShipVia=""
gcSycd = "0"  && nq or another station
glFut = .f.  &&future ship 
gcService = ""
PUBLIC LCSTR, glProc     
glProc = .f.		&& has this been processed

gnMultCOD = 0
*** used for filelink interface
public glOKtoship 
glOKtoship = .t.

***trade=STATION 1 or 2;
** trumble=STATION 3 or 4
***GCSYCD = '1' IS STILL TRADE    
** GCSYCD = '2' IS STILL TRUMBLE 
**
if s_usernum = "11" or s_usernum = "12"  &&trade
	gcsycd = "1"
	WAIT WINDOW " STATION 1 OR 2 IS TRADE" TIMEOUT .5
	 *MESSAGEBOX("Logged into SPSS TEST SYSTEM FOR TESTING IN R3Q" + CHR(13),0+48,"SAP R3Q TESTING")
	 
endif

** trumble consumer is
*if s_usernum = "13" or s_usernum = "15"	&& added 15 per tom 6/20/2002 
	if s_usernum = "13" or s_usernum = "15"	&& 
	gcsycd = "2"
	WAIT WINDOW " TRUMBLE STATION 3 or 5 " TIMEOUT .5
	*MESSAGEBOX("Logged into SPSS TEST SYSTEM FOR TESTING IN R3Q" + CHR(13),0+48,"SAP R3Q TESTING")
endif

   if s_usernum = "14" &&mobile station cna be either trade or trumble
   lnanswer = MESSAGEBOX("Trade/Consumer Station" + CHR(13)+ "Do you want to log into Trade?",4+32,"Station 4 Trade or Consumer?")
     IF lnanswer = 6	&& yes
       gcsycd = "1"
       MESSAGEBOX("Logged in as TRADE" + CHR(13),1+48,"Trade or Consumer?")
     *  MESSAGEBOX("Logged into SPSS TEST SYSTEM FOR TESTING IN R3Q" + CHR(13),0+48,"SAP R3Q TESTING")
	        else
	gcsycd = "2"
	messagebox ( "CONSUMER Station on 4 " + CHR(13),1+48,"Trade or Consumer?")
	*MESSAGEBOX("Logged into SPSS TEST SYSTEM FOR TESTING IN R3Q" + CHR(13),0+48,"SAP R3Q TESTING")
	 
	endif
endif

if gcsycd <> "0"	&& NQ or another station

	IF type("gnSqlsap")="U"	&& not defined
		public gnsqlsap
		gnSqlSAP =SQLCONNECT('SAPODBC','apssODBC','apssODBC')
		IF GNSQLSAP >0 
    		ELSE	
			MESSAGEBOX("NOT CONNECTED TO SAP")
		ENDIF
	endif
endif		&& syscd <>"0"





****************************************************************************
* CUS_MISC() function will be called when the following events occurs,
* you can enable or disable execution of CUS_MISC() by setting
* elements of array SC_OPTN.SA_POWERACT[n] to .T. or .F.
* whrere n can be :
*    _PWA_BEF_VOID                before package void
*	    _PWA_AFT_VOID                after package void
*    _PWA_BEF_UNVOID              before package unvoid
*    _PWA_AFT_UNVOID              after package unvoid
*    _PWA_BEF_REPORT              before printing a report
*    _PWA_AFT_REPORT              after printing a report
*    _PWA_BEF_MANIFEST            before manifest print
*    _PWA_AFT_MANIFEST            after manifest print
*    _PWA_BEF_DAYOPEN             before day open
*    _PWA_AFT_DAYOPEN             after day open
*    _PWA_BEF_DAYCLOSE            before day close
*	    _PWA_AFT_DAYCLOSE            after day close
*    _PWA_BEF_SWOG                before SWOG screen
*    _PWA_AFT_SWOG                after SWOG screen
*    _PWA_GOTOSHIP                when starting the shipping screen
*    _PWA_GOTOVIEW                when starting the view manifest screen
*    _PWA_SORTBATCH               Build index for batchlink custom access
*    _PWA_GETBATCHREC             Select batchlink record to process return .F.
*    _PWA_ERRORHANDLE             error handle
*    _PWA_MISCFEE                 misc fee calculation
*    _PWA_REINDEX                 rebuild custom index files
*    _PWA_ADDSWOGLINE             add SWOG line
*    _PWA_AFT_ALTRATE             after alternate rate calculation
*    _PWA_AFT_SHOPRATE            after calculating rate in shopping process
*    _PWA_AFT_THERMAL_LBL         after sending label data to a thermal printer
*                                 the call will happen only if the post sequence
*                                 of the virtual printer is not blank
*    _PWA_AFT_FEXPFRG_LBL         after sending label data to Unimark/Datamax
*                                 thermal printer for FedEx label.
*                                 The call will happen only prior to
*                                 to printing the barcode. This is needed
*                                 to vercome a printer problem.
*    _PWA_GETHOLDREC              custom retreive from hold
*
*    _PWA_MAILING_LABEL           Insert your code to replace the standard
*                                 Ascent mailing label print code
*    _PWA_SHIPPING_LABEL          Insert your code to replace the standard
*                                 Ascent shipping label print code
*    _PWA_SMART_DIT               Returns delivery commitment
*    _PWA_SMART_DIT               Returns delivery commitment
*    _PWA_SHOPSCREEN              Do not use
*    _PWA_SHOPSCREEN              Do not use
*    _PWA_FUTURE_CLICK            After the click event of future ship command button

With Sc_Optn
*  .Sa_PowerAct[_PWA_BEF_VOID]        = .T.    && before package void
  .Sa_PowerAct[_PWA_AFT_VOID]        = .T.    && after package void
*  .Sa_PowerAct[_PWA_BEF_UNVOID]      = .T.    && before package unvoid
*  .Sa_PowerAct[_PWA_AFT_UNVOID]      = .T.    && after package unvoid
*  .Sa_PowerAct[_PWA_BEF_REPORT]      = .T.    && before printing a report
*  .Sa_PowerAct[_PWA_AFT_REPORT]      = .T.    && after printing a report
*  .Sa_PowerAct[_PWA_BEF_MANIFEST]    = .T.    && before manifest print
*  .Sa_PowerAct[_PWA_AFT_MANIFEST]    = .T.    && after manifest print
*WAIT WINDOW "THIS IS BEFORE OPEN DAY" TIMEOUT 2
*  .Sa_PowerAct[_PWA_BEF_DAYOPEN]     = .T.    && before day open
 .Sa_PowerAct[_PWA_AFT_DAYOPEN]     = .T.    && after day open
*wait window " AFTER DAY OPEN CODE HERE" TIMEOUT 3.5

*  .Sa_PowerAct[_PWA_BEF_DAYCLOSE]    = .T.    && before day close
  .Sa_PowerAct[_PWA_AFT_DAYCLOSE]    = .T.    && after day close
*  .Sa_PowerAct[_PWA_BEF_SWOG]        = .T.    && after SWOG screen
*  .Sa_PowerAct[_PWA_AFT_SWOG]        = .T.    && after SWOG screen
*  .Sa_PowerAct[_PWA_GOTOSHIP]        = .T.    && when starting the shipping screen
*  .Sa_PowerAct[_PWA_GOTOVIEW]        = .T.    && when starting the view manifest screen
*  .Sa_PowerAct[_PWA_SORTBATCH]       = .T.    && Build index for batchlink custom access
*  .Sa_PowerAct[_PWA_GETBATCHREC]     = .T.    && Select batchlink record to process return .F.
*  .Sa_PowerAct[_PWA_ERRORHANDLE]     = .T.    && Error Handler
*  .Sa_PowerAct[_PWA_MISCFEE]         = .T.    && Used Defined Misc Fee
*  .Sa_PowerAct[_PWA_REINDEX]         = .T.    && Used Defined Pack and Reindex Routine
*  .Sa_PowerAct[_PWA_ADDSWOGLINE]     = .T.    && add SWOG line
*  .Sa_PowerAct[_PWA_AFT_ALTRATE]     = .T.    && after alternate rate calculation
*  .Sa_PowerAct[_PWA_AFT_SHOPRATE]    = .T.    && after calculating rate in shopping process
*  .Sa_PowerAct[_PWA_GETHOLDREC]      = .T.    && custom retreive from hold
  .Sa_PowerAct[_PWA_AFT_THERMAL_LBL] = .T.    && after sending label data to a thermal printer
*                                              && the call will happen only if the post sequence
*                                              && of the virtual printer is not blank
*  .Sa_PowerAct[_PWA_AFT_FEXPFRG_LBL] = .T.    && after sending label data to Unimark/Datamax
                                               && thermal printer for FedEx label.
                                               && The call will happen only prior to
                                               && to printing the barcode. This is needed
                                               && to vercome a printer problem.
*  .Sa_PowerAct[_PWA_MAILING_LABEL]   = .T.    && Insert your code to replace the standard
                                               && Ascent mailing label print code
*  .Sa_PowerAct[_PWA_SHIPPING_LABEL]  = .T.    && Insert your code to replace the standard
                                               && Ascent shipping label print code
*  .Sa_PowerAct[_PWA_SMART_DIT]       = .T.    && Value route delivery commitment
*  .Sa_PowerAct[_PWA_SHOPSCREEN]      = .T.
*  .Sa_PowerAct[_PWA_FUTURE_CLICK]    = .T.    &&After the click event of futureship command button
EndWith
Return .T.
EndFunc

****************************************************************************
FUNCTION Cus_SetInit()
****************************************************************************
* Add here code to be executed after Login Screen during APSS-SET Entry
* Return .F. to Cancel ASC-SET Execution
*
* To add you pack and reindex code to APSS-SET amke sure that you do one
* of the following:
* Uncomment the next line

* Sc_Optn.Sa_PowerAct[_PWA_REINDEX]  = .T.           && User Defined Pack and Reindex Routine

* or uncomment the next line to activate the Cus_Init function if
* it is OK to share the APSS initialization function with APSS-SET intialization.

***Return Cus_Init()

Return .T.
EndFunc

****************************************************************************
FUNCTION Cus_RsvInit()
****************************************************************************
* Add here code to be executed after successful login to Ascent-OLE
* Return .F. to Cancel Ascent-OLE login
***Return Cus_Init()
Return .T.
EndFunc

****************************************************************************
FUNCTION Cus_AdmInit()
****************************************************************************
* Add here code to be executed after successful login to APSS_ADM
* Return .F. to Cancel ASC_ADM login
***Return Cus_Init()
Return .T.
EndFunc

****************************************************************************
FUNCTION Cus_End()
****************************************************************************
* Add here code to be executed after the user select to exit Ascent
goDIT = .F.
Return .T.
EndFunc

****************************************************************************
FUNCTION Cus_SetEnd()
****************************************************************************
* Add here code to be executed after the user select to exit Asc-set
***Return Cus_End()
Return .T.
EndFunc

****************************************************************************
FUNCTION Cus_RsvEnd()
****************************************************************************
* Add here code to be executed when the logout method in executed
***Return Cus_End()
Return .T.
EndFunc

****************************************************************************
FUNCTION Cus_AdmEnd()
****************************************************************************
* Add here code to be executed when the logout method in executed
***Return Cus_End()
Return .T.
EndFunc

****************************************************************************
FUNCTION Cus_Rate()
****************************************************************************
* Add here code to be executed for user defined rating method
* Not implemented  -- do not use!!!!
Return .T.
EndFunc

****************************************************************************
FUNCTION Cus_AltRate()
****************************************************************************
* Add here code to be executed for user defined alternate rate method
* Not implemented  -- do not use!!!!
Return .T.
EndFunc

****************************************************************************
FUNCTION Cus_Misc(tnPowerAct,tParm1,tParm2,tParm3,tParm4)
****************************************************************************
* Misc poweract action points
*
* tnPowerAct 		= Point
* tParm1 - tParm4	= Misc parameters varies per point
*
* Note that CUS_MISC will not be called if an activity was initiated by the
* task scheduler. (You can activate it as a VFP function from the Task Scheduler)
*
DO CASE
	CASE tnPowerAct = _PWA_BEF_VOID 	&& before package void
*    tParm1 - tParm4 not used, current table will be with alias MAN and current
*              record will be the record to be voided.
		Return .T.

	CASE tnPowerAct = _PWA_AFT_VOID 	&&  after package void
*    tParm1 - tParm4 not used, current table will be with alias MAN and current
*              record will be the record to be voided.
	
*****void code********
**
*
LCFLAG = ""
STORE ALLT(MAN.MSODBC) TO LCfLAG 

IF LEN(ALLT(LCFLAG)) > 0  && ONLY RUN VOID CODE FOR INTERFACED SHIPMENTS
**** sap void
	IF alltr(LCFLAG)== "S" or alltr(substr(lcFlag,1,3))="SDF"		&& SAP INTERFACED SHIPMENT
	WAIT WINDOW "INDICATOR =  "+ALLTR(LCFLAG) TIMEOUT 1
		IF gnSQLSAP > 0
			STORE allt(MAN.UNIQUEID) TO lcID
			STORE ALLT(MAN.PKGID) TO LCPKG

** TRUMBLE VOID ONLY- THERE IS NO VOID FOR TRADE. FOR TRADE, WE PUT IN A 'V' IN THE INSERT STATEMENT IF THE SHIPMENT HAS ALREADY BEEN PROCESSED AND VOIDED.
** ALL SAP SHIPMENTS FOR AN ORDER MUST BE VOIDED IN ORDER TO PROCEED
			IF GCSYCD = "2"	&& TRUMBLE VOID
				WAIT WINDOW "VOIDING SAP TRUMBLE ODBC SHIPMENT " + LCID+ " PACKAGE " +LCPKG TIMEOUT .5
				XX = SQLEXEC(gnsqlSAP, "delete from PRD.prd.ZAPSS where PRD.prd.ZAPSS.MNUNID = ?lcid")
					if xx>0 
						wait window " VOID WAS SUCCESSFUL" TIMEOUT .5
					else
						wait window " VOID FAILED!!" TIMEOUT 1
					endif
			ENDIF	&& TRUMBLE VOID** end of trumble void		
			ELSE	&& NOT CONNECTED
				MESSAGEBOX("CANNOT VOID.. SAP ODBC CONNECTION IS NOT PRESENT")
		ENDIF	&& SQLHANDLE IS NOT CONNECTED
	ENDIF

ELSE && not an interfaces shipment

	WAIT WINDOW " MANUAL SHIPMENT VOID" TIMEOUT 1
ENDIF	&& NOT AN INTERFACED SHIPMENT
	
		Return .T.

	CASE tnPowerAct = _PWA_BEF_UNVOID 	&& before package unvoid
*    tParm1 - tParm4 not used, current table will be with alias MAN and current
*              record will be the record to be unvoided.
		Return .T.

	CASE tnPowerAct = _PWA_AFT_UNVOID 	&& after package unvoid
*    tParm1 - tParm4 not used, current table will be with alias MAN and current
*              record will be the record to be unvoided.
		Return .T.


	CASE tnPowerAct = _PWA_BEF_REPORT 	&& before printing a report
*    tParm1 = Report #  (see c_cust.def)
*    tParm2 - tParm4 not used
		Return .T.

	CASE tnPowerAct = _PWA_AFT_REPORT 	&& after printing a report
*    tParm1 = Report # (see c_cust.def)
*    tParm2 - tParm4 not used
		Return .T.

	CASE tnPowerAct = _PWA_BEF_MANIFEST && before manifest print
*    tParm1 = Carrier Code
*    tParm2 - tParm4 not used
		Return .T.

	CASE tnPowerAct = _PWA_AFT_MANIFEST && after manifest print
*    tParm1 = Carrier Code
*    tParm2 - tParm4 not used
		Return .T.

	CASE tnPowerAct = _PWA_BEF_DAYOPEN 	&& before day open
*    tParm1 = Carrier Code
*    tParm2 - tParm4 not used
		Return .T.

	CASE tnPowerAct = _PWA_AFT_DAYOPEN 	&& after day open
*    tParm1 = Carrier Code
*    tParm2 - tParm4 not used

	AFTER_FSMOVE()
		Return .T.

	CASE tnPowerAct = _PWA_BEF_DAYCLOSE && before day close
*    tParm1 = Carrier Code
*    tParm2 - tParm4 not used
		Return .T.

	CASE tnPowerAct = _PWA_AFT_DAYCLOSE && after day close
*    tParm1 = Carrier Code
*    tParm2 - tParm4 not used
	IF tParm1 = "PR8"  &&take out if statement & change functionappend
		FUNCTIONAPPEND()
		ENDIF  
		Return .T.

	CASE tnPowerAct = _PWA_BEF_SWOG 	&& Before SWOG Popup
*    tParm1 - tParm4 not used
		
		

	CASE tnPowerAct = _PWA_AFT_SWOG 	&& After SWOG Popup
*    tParm1 - tParm4 not used
		Return .T.

	CASE tnPowerAct = _PWA_GOTOSHIP 	&& when starting the shipping screen
*    tParm1 - tParm4 not used
		Return .T.

	CASE tnPowerAct = _PWA_GOTOVIEW 	&& when starting the view manifest screen
*    tParm1 - tParm4 not used
		Return .T.

	CASE tnPowerAct = _PWA_SORTBATCH 	&& build batchlnk file index
*    tParm1 - index file name
*    tParm2 - tParm4 not used
		Return .T.

	CASE tnPowerAct = _PWA_GETBATCHREC 	&& Select batchlink record to process
*    tParm1 - tParm4 not used
*              record will be the record to be processed.
		Return .T.

	CASE tnPowerAct = _PWA_ERRORHANDLE 	&& error handler
*    tParm1 - error #
*    tParm2 - error message
*    Return .T. to proceed with default Ascent error handling
*    Return .F. - does not display Ascent error message
*                 in that case if you set SC_SHIP.CANCELPACKAGE to .T. the
*                 current package processing will terminate, if this is box #1
*                 and data came from an interface, the interface loop will
*                 repeat. If you intend to have a shipping session for hands free
*                 operation set the property SF_SHIP.HANDSFREE to .T. in that
*                 case error messages that direct user to some activities will
*                 not pop up a window and a default will be taken by the system
*                 unless you set SC_SHIP.CANCELPACKAGE to .T.
		Return .T.

	CASE tnPowerAct = _PWA_MISCFEE 		&& User Defined Misc Fee
*    tParm1 - Logical - the state of the misc fee checkbox .T. - ON, .F.-OFF
*    tParm2 - Carrier
*    tParm3 - Service
*    tParm4 - not in use
*    The function sould populate up to 6 fees in the array
*    SC_SHIP.A_MISCFEE[n,m]
*    Where n is a number 1 to 6
*          m = 1 = misc fee description
*          m = 2 = misc fee amount
*          m = 3  .T. - misc fee is on, .F. - misc fee is off
		Return .T.

	CASE tnPowerAct = _PWA_REINDEX 		&& User Defined Pack and Reindex
*    tParm1 - tParm4 not used
		Return .T.

        CASE tnPowerAct = _PWA_ADDSWOGLINE    && after click of ADD command button of the SWOG screen
*    tParm1 - tParm4 not used
*    the SWOG line will not be added if the function returns .F.
		Return .T.

	CASE tnPowerAct = _PWA_AFT_ALTRATE      && After alternate rate calculation
*    tParm1 - tParm4 not used
*    This function will be called after Ascent calculates the alternate
*    rates for a package.
*    At this point you will have 3 fees available. You can use this rates
*    to determine the value you want to place in the alternate base fee.
*
*    SC_SHIP.A_SHIP[_SHP_FEE] - consolidated base fee - note that
*                               Ascent will consolidate only if the consolidtion
*                               is turned on for the carrier, the shipping
*                               mode is set to consolidate upfront and
*                               the multipkg shipment qualifies for
*                               consolidation (as Ascent sees it).
*
*    SC_SHIP.A_SHIP[_SHP_NORMCHG] - non consolidated fee
*
*    SC_SHIP.A_SHIP[_SHP_AFEE] - Alternate rate base fee as calculated by
*                                Ascent (based on alternate method set for
*                                the carrier, if it is set to None the
*                                fee equals to SC_SHIP.A_SHIP[_SHP_FEE].
*
*    Your code will calculate the alternate rates fee and save it into
*    SC_SHIP.A_SHIP[_SHP_AFEE].
*
*    Note that Ascent will add to the base alt fee the special services
*    alternate fee prior to saving it in the manifest element (ALTFEE).
		Return .T.

	CASE tnPowerAct = _PWA_AFT_SHOPRATE     && After calculating rate in shopping process
*    tParm1 = Numeric   Rate - Ascent calculated shopping rate for all boxes
*    tParm2 = Logical   Consolidation state
*              .F. - tParm1 is the non consolidated base fee for the boxes
*              .T. - tParm1 is the consolidated base fee for the boxes
*    tParm3 = Character  Shopping Key
*
*    This function will be called up to two times for each carrier/service
*    which is a member of the selected shopping key.
*    The first time it is called with the total nonconsolidated base
*    charges (tParm2 is .F.) in your code you can manipulate tParm1 and return
*    it to Ascent.
*
*    The second time it is called with the total consolidated base charges
*    (tParm2 is .T.). Only if the carrier is set to shop on consolidated
*    charges, and the shipment qualifies for consolidation from Ascent point
*    of view. In your code you can manipulate tParm1 and return it to Ascent, if
*    you do not want to consider the consolidated rates for a given
*    shipment, set tParm1 to 9999999.
		Return tParm1

	CASE tnPowerAct = _PWA_AFT_THERMAL_LBL	&& called at the end of thermal label printout,
*    tParm1 = Numeric - what is printed
*             1 = shipping label
*             4 = audit label
*             6 = custom form
*    the return value from this call will be sent to the thermal printer.
*    Note that this function will execute only if the post sequence for
*    the selected thermal printer is non blank.
*    Your code will be called with 1 parameter
	local lcstr
		lcstr = ANSI_CR_LF 
	
*!*		DO CASE
*!*			case P_1 = 1 

	IF substr(s_shipper,1,3) = "POS" or substr(s_shipper,1,4) = "UPS1"
		lcanytext = ""
		store allt(man.anytext2) to lcAnytext 
	
		if LEN(ALLT(lcAnytext)) >0
			lcstr = lcstr +"^FT25,1250^A0N,50,45^FDSent to you " + lcanytext +"^FS" + ANSI_CR_LF 
		ENDIF
			endif
	
	IF SUBSTR(S_SHIPPER,1,3)="POM"
*!*			IF SUBSTR(L_PKID,1,10)=ALLTRIM(L_INV) &&SOME NON-INTERFACED ORDERS
*!*			LCSTR=LCSTR + "^FT0447,0840^A0,30,24   ^FR^FDPKG ID " + L_PKID + "^FS" + ANSI_CR_LF
*!*		
*!*	  		  ELSE    
*!*	        	 LCSTR=LCSTR + "^FT0447,0840^A0,30,24   ^FR^FDPKG ID " + L_INV + "^FS" + ANSI_CR_LF 
*!*	   		 ENDIF
*WAIT WINDOW "POSTAL MANIFEST"
 ****IF JUST NEED IT ON BOUND PRINTED MATTER UNCOMMENT THE NEXT LINE DOWN
*!*		gcService = ALLTRIM(man.upstyp)
*!*			IF ALLTRIM(gcService) ="4T"
*!*			WAIT WINDOW "BOUND PRINTED MATTER " + ALLTRIM(gcService)
*!*				LCSTR=LCSTR + "^FT0041,0450^A0,30,24   ^FR^FDRETURN SERVICE REQUESTED" + "^FS" + ANSI_CR_LF
*!*			ELSE
*!*			ENDIF
	LCSTR=LCSTR + "^FT0041,0450^A0,30,24   ^FR^FDRETURN SERVICE REQUESTED" + "^FS" + ANSI_CR_LF
	ELSE
	ENDIF	
*** PRINTER POST SEQUENCE MUST BE ^XZ
    RETURN LCSTR
*!*	ENDCASE
		Return ""
        CASE tnPowerAct = _PWA_AFT_FEXPFRG_LBL
           && called at the end of FEXP Unimark/Datamax thermal label printout,
           && before the barcode printing
*    tParm1 = Numeric - what is printed
*             1 = shipping label
*    the return value from this call will be sent to the thermal printer.
*    Note that this function will execute only if the post sequence for
*    the selected thermal printer is non blank.
*    Your code will be called with 1 parameter
                Return ""

	CASE tnPowerAct = _PWA_GETHOLDREC       && Custom retrieve from hold
*    tParm1 - tParm4 not used
*    Use this function to select a record from the pack & hold file, note
*    the file will be opened with an alias of BATCH and the index order
*    set to 2 (consignee name). If your function returns .T. the selected
*    record will be processed if it returns .F. user will leave retrieve
*    from hold mode.
*    Record will be the record to be processed.
		Return .T.

        CASE tnPowerAct = _PWA_MAILING_LABEL
*Insert your code to replace the standard Ascent mailing label print code
*Return empty string to prevent the standard Ascent mailing label from being printed
*Return non empty string, Ascent will send the string to the printer and the
*standard Ascent mailing label will not print.
*
*The mailing address is availabel in  SC_SHIP.A_CUS as follows:
*  SC_SHIP.A_CUS[_C_CSNAM1]  - comapny
*  SC_SHIP.A_CUS[_C_CSNAM2]  - attention
*  SC_SHIP.A_CUS[_C_CSADD1]  - address
*  SC_SHIP.A_CUS[_C_CSADD2]
*  SC_SHIP.A_CUS[_C_CSADD3]
*  SC_SHIP.A_CUS[_C_CSCITY]  - city
*  SC_SHIP.A_CUS[_C_CSSTAT]  - state
*  SC_SHIP.A_CUS[_C_CSZIPC]  - ZIP
        RETURN("")

        CASE tnPowerAct = _PWA_SHIPPING_LABEL
*Insert your code to replace the standard Ascent shipping label print code
*Return empty string to prevent the standard Ascent shipping label from being printed
*Return non empty string, Ascent will send the string to the printer and the
*standard Ascent shipping label will not print.
*
*    tParm1 - Logical  .F. - print shipping label
*                      .T. - used only for FedEx power to designate
*                            printout of return COD thermal label
        RETURN("")

       CASE tnPowerAct = _PWA_SMART_DIT  &&Return delivery commitment code
*    tParm1 = Carrier Code
*    tParm2 - tParm4 not used
        RETURN(99)

      CASE _PWA_SHOPSCREEN
         RETURN(.T.)

      CASE tnPowerAct = _PWA_FUTURE_CLICK
* Insert your code to perform and required activities when future ship
* mode is being toggeled.
         RETURN(.T.)
ENDCASE
Return .T.
EndFunc

****************************************************************************
FUNCTION Cus_CheckDigit(tcTrackingNumber)
****************************************************************************
* Put here code to be calculate check digit for the tracking number
* and return a complete tracking # by substituting the last digit with
* the calculated checkdigit.
* Note that tcTrackingNumber includes checkdigits which may be incorrect,
* this function should calculate the right one return a complete
* tracking #.
Return tcTrackingNumber
EndFunc

****************************************************************************
FUNCTION Cus_Rsv(tParm1)
****************************************************************************
*put here code to be executed from the custom method of the Ascent OLE server
*tParm1 - one parameter that you can pass to the method
*
*note that the Ascent OLE control object is in a public variable SF_RSV
*
Return .T.
EndFunc

****************************************************************************
FUNCTION Cus_Adm(tParm1)
****************************************************************************
*put here code to be executed from the custom method of the APSS ADM server
*tParm1 - one parameter that you can pass to the method
*
*note that the ASC_ADM control object is in a public variable SF_RSV
*
Return .T.
EndFunc

****************************************************************************
PROCEDURE  CUS_SETRETADDR(P_CSNAM,P_CSADD1,P_CSADD2,P_CITY,P_PHONE)
****************************************************************************
* The return Address is put here and can be manipulated
* before the monarch prints it on standard labels
*
* SUBSTR(Setup.Flags,4,1) == 'M' must be true for this to work
*

ENDPROC

****************************************************************************
PROCEDURE usptrack	&& populates tracking # for POM/POS
****************************************************************************
STORE "" TO lcUsp_Man
STORE "" TO lc_uspnum
*** determine which manifest is being used
*
DO CASE
	CASE ALLTRIM(SC_SHIP.A_SHIP[79]) = "POS"
		lcUsp_Man = "POS"
	CASE ALLTRIM(SC_SHIP.A_SHIP[79]) = "POM"
		lcUsp_Man = "POM"
	OTHERWISE
		lcUsp_Man = ""
ENDCASE

*** only process for USP or USP1
*
IF ! EMPTY(lcUsp_Man)

STORE SELECT() TO lnOldSele
*** increment counter in file
*
	USE usptrack IN 0
	SELE usptrack
	GO TOP
	STORE usptrack.usptrack + 1 TO ln_uspnum
	REPLACE usptrack.usptrack WITH ln_uspnum
*** create unique tracking number
*
IF lcUsp_Man = "POS"
	STORE "POS"+ALLTRIM(STR(ln_uspnum)) TO lc_uspnum
*	wait window "Updating USP Number "+lc_uspnum TIMEOUT 3
	IF USED('uspman')
		SELE(uspman)
		REPLACE uspman.trknum WITH lc_uspnum
			ELSE
		*wait window "File POSman.dbf not found" TIMEOUT 2 
	ENDIF
		ELSE
	STORE "POM"+ALLTRIM(STR(ln_uspnum)) TO lc_uspnum
	*wait window "Updating POM Number "+lc_uspnum TIMEOUT 3
	
	IF USED('usp1man')
		SELE(usp1man)
		REPLACE usp1man.trknum WITH lc_uspnum
	ENDIF
			
ENDIF
*** back to original work area
*
SELE(lnOldSele)

*** update MAN.trknum
*
IF USED('man')
	REPLACE man.trknum WITH lc_uspnum
		ELSE
	wait window "MAN not found" TIMEOUT 2
ENDIF

*** cleanup code
*
IF used('usptrack')
	USE IN usptrack
ENDIF

ENDIF	&& main condition

RELEASE lnOldSele, ln_uspnum, lcUsp_Man, lc_uspnum

RETURN

*****
**		COD_OPTIONS  FROM LEANINTREE   used in after post
*************************************************************************************
PROCEDURE cod_options	&& after post
*************************************************************************************
**CALLING THIS AFTER POST BUT EVERYTHING IS IN THE ARRAY
**PROCEDURE VS FUNCTION??

*WAIT WINDOW "STARTING COD OPTIONS " TIMEOUT 2 &&does this
** used because maxzicode Takes away ZIP+4 from ZIP  added 10/20/2002
IF SUBSTR(S_SHIPPER,1,3) = "UPS"	
	LCZIP = SUBSTR(ALLT(SC_SHIP.A_CUS(8)),1,5)
	STORE LCZIP TO SC_SHIP.A_CUS(8)
ENDIF

  IF ALLT(SC_SHIP.A_SHIP(66))="Y" AND  SC_SHIP.A_SHIP(24)= SC_SHIP.A_SHIP(25)

	 Sele * From man Into Cursor our_man Where ALLTRIM(SUBSTR(pkgid,1,10)) = Allt(gcorder) And SHPSTS <> "V"
		Sele our_man
		LNTOTFRT=0
		LNTOTFEE=0
        Calc Sum(our_man.FRIGHT) To LNTOTFRT 
	    Calc Sum(our_man.CODFEE) To LNTOTFEE 
	DO CASE 

		CASE sc_ship.a_ship[31] = 0  &&changed 10/24/05 per tom-works
			LNHAND=0
			*lnhand = (SC_SHIP.A_SHIP[3]*.045)
			*WAIT WINDOW "freight=  "=STR(lntotfrt,5,2)
			*WAIT WINDOW "HAND CHG= "+STR(LNHAND,5,2)
          *  sc_ship.a_ship[16] = sc_ship.a_ship[16]+ LNTOTFRT + LNHAND  &&TOOK OUT HANDLING CHARGE 7/2008
			sc_ship.a_ship[16] = sc_ship.a_ship[16]+ LNTOTFRT 
            gnMultCOD = sc_ship.a_ship[16]
            
		CASE sc_ship.a_ship[31] = 1  
			sc_ship.a_ship[16] = sc_ship.a_ship[15]
					
		CASE sc_ship.a_ship[31] = 2  &&changed 10/24/05 per tom-works
			sc_ship.a_ship[16] = sc_ship.a_ship[15]+sc_ship.a_ship[17]
				
	ENDCASE
endif	&& not a cod shipment

RETURN

*****************************************************************************
function e_pack()
*****************************************************************************
******* end of package FOR TRADE ONLY!!!!!!!
***** populate variables to pass in the 'header' INSERT statement ***
***   "S" IS FOR SAP SHIPMENTS
***adds in handling fee if a prepaid type "1"

IF SC_SHIP.A_PRV(59)<>"Y"  &&not a future shipment
		
	IF ! EMPTY(SA_CUST(5))
   	 replace man.msodbc with sa_cust(5)		&& flag for odbc voids
    ENDIF
    
	IF ! EMPTY(SA_CUST(8))
	   REPLACE MAN.uccstornum WITH SA_CUST(8)
    ENDIF
	
		ELSE

		glFut = .T. &&IS a Future Shipment
    	
    	IF ! EMPTY(SA_CUST(5))
		    replace man.msodbc with sa_cust(5) + "FS"		&& flag for odbc + INDICATE IT'S A FUTURE SHIP
    	ENDIF
             
		IF ! EMPTY(SA_CUST(8))
   			REPLACE MAN.uccstornum WITH SA_CUST(8)	
    	ENDIF
    
ENDIF   &&is this a future ship?   
	lnbase = 0
	lnbase = man.shipchg + man.genchg1
	*WAIT WINDOW "base charge=  "+ STR(lnbase,5,2)
	lntotfrt = 0
	lnhand = 0
		
		IF man.Payment =1 AND MAN.ANYNUM=0.00 AND GCSYCD == "1" 
				*lnhand = (lnbase*.045) &&CHANGED TO ZERO BUT LEAVE CODE IN PLACE
			*	wait WINDOW "Prepay and Trade side"  
			lnhand = 0
			*lntotfrt=lnbase+lnhand
			lntotfrt=lnbase
			LNCOD=MAN.COD
			LNCODFEE =MAN.CODFEE
			LNCODCHG=LNTOTFRT+LNCOD+LNCODFEE
			
		
		replace man.shipchg WITH lntotfrt &&incl handling fee with out fee exposure per Tom 070607
		replace man.fright WITH lntotfrt  &&incl handling fee with out fee exposure
		replace man.altfee WITH lntotfrt  &&incl handling fee with out fee exposure
		
			IF man.codfee=9.00
					REPLACE MAN.CODCHG WITH LNCODCHG
					replace man.acctchg WITH gnMultCOD
				ELSE
			ENDIF
		
	*	WAIT WINDOW "FREIGHT= "+STR(LNTOTFRT,5,2)
	*	WAIT WINDOW "COD= "+STR(LNCOD,5,2)
	*	WAIT WINDOW "COD,FREIGHT + FEE= "+STR(LNCODCHG,5,2)
		*WAIT WINDOW "total freight=  "+STR(lntotfrt,5,2)
				ELSE
				*	STORE lnbase TO lntotfrt
				ENDIF
	**clear out gnMultCOD
	gnMultCOD = 0

RETURN		&& return_order()

****************************************************************************
function e_packtrum()
*****************************************************************************
******* end of shipment for TRUMBLE SHIPMENTS ONLY!!!!!
***** populate variables to pass in the 'header' INSERT statement ***
***   "S" IS FOR SAP SHIPMENTS
***adds in handling fee if a prepaid type "1"
IF GCSYCD = "2"
	REPLACE MAN.USERFLD3 WITH STR(SA_CUST(3),7,2)
	*WAIT WINDOW "USERFLD SHOUDL BE  "+ VAL(STR(SA_CUST(3)))
	ENDIF
	
	

IF SC_SHIP.A_PRV(59)<>"Y"  &&not a future shipment
		
	IF ! EMPTY(SA_CUST(5))
   	 replace man.msodbc with sa_cust(5)		&& flag for odbc voids
    ENDIF
    
	IF ! EMPTY(SA_CUST(8))
	   REPLACE MAN.uccstornum WITH SA_CUST(8)
    ENDIF
	
		ELSE

		glFut = .T. &&IS a Future Shipment
    	
    	IF ! EMPTY(SA_CUST(5))
		    replace man.msodbc with sa_cust(5) + "FS"		&& flag for odbc + INDICATE IT'S A FUTURE SHIP
    	ENDIF
             
		IF ! EMPTY(SA_CUST(8))
   			REPLACE MAN.uccstornum WITH SA_CUST(8)	
    	ENDIF
    
ENDIF   &&is this a future ship?   
	lnbase = 0
	lnbase = man.shipchg + man.genchg1
	*WAIT WINDOW "base charge=  "+ STR(lnbase,5,2)
	lntotfrt = 0
	lnhand = 0
		
		IF man.Payment =1 AND MAN.ANYNUM=0.00 AND GCSYCD == "1" 
				lnhand = (lnbase*.045)
			*	wait WINDOW "Prepay and Trade side"  
			lntotfrt=lnbase+lnhand
			LNCOD=MAN.COD
			LNCODFEE =MAN.CODFEE
			LNCODCHG=LNTOTFRT+LNCOD+LNCODFEE
			
		
		replace man.shipchg WITH lntotfrt &&incl handling fee with out fee exposure per Tom 070607
		replace man.fright WITH lntotfrt  &&incl handling fee with out fee exposure
		replace man.altfee WITH lntotfrt  &&incl handling fee with out fee exposure
		
			IF man.codfee=9.00
					REPLACE MAN.CODCHG WITH LNCODCHG
				ELSE
				replace man.userfld3 WITH STR(sa_cust(3))
				
			ENDIF
		
	*	WAIT WINDOW "FREIGHT= "+STR(LNTOTFRT,5,2)
	*	WAIT WINDOW "COD= "+STR(LNCOD,5,2)
	*	WAIT WINDOW "COD,FREIGHT + FEE= "+STR(LNCODCHG,5,2)
		*WAIT WINDOW "total freight=  "+STR(lntotfrt,5,2)
				ELSE
				*	STORE lnbase TO lntotfrt
				ENDIF


RETURN		&& return_order()
*

**************************************************************************
FUNCTION Ship_via()		&& use this to PREPOPULATE using translated fields
**************************************************************************

lnOldSele=SELECT()
lnOldRec=RECNO()

USE scntrn IN 0 ALIAS trans AGAIN
SELE trans
GO TOP

LOCATE FOR ALLT(scntrcode) == ALLT(Sc_Ship.l_trncode)

IF FOUND()
	Sc_Ship.a_ship(77)=ALLT(trans.service)
	Sc_Ship.a_ship(79)=ALLT(trans.carrier)
else
	Sc_Ship.a_ship(77)= "G"
	Sc_Ship.a_ship(79)="UPS"    &&lp commented out 06/04
	s_shipper = "UPS"            &&changed to RPS for new default carrier
ENDIF

if ALLT(Sc_Ship.l_trncode) == "C"
	sc_ship.a_cus(9) = "CA"	&& canadian shipment		
endif 

IF USED('trans')
	USE IN trans
ENDIF

SELECT (lnOldSele)
RECNO(lnOldRec)
		
RETURN 

****************************************************************************
Function e_ship()
***************************************************************************
sc_ship.a_prv(71) =.f.  &&turns off multiple packages
sc_ship.a_ship(51) =.f.  &&turns off multiple packages
****
lnoldsele = Select()

lcCurrinv = ALLT(man.invnum)
IF glFut = .f.
	SELE * FROM MAN WHERE atcc(lccurrinv, invnum) >0 and man.shpsts <> "V" into cursor our_man
	SELE OUR_MAN

	GO TOP
	SCAN

	IF LEN(ALLT(OUR_MAN.MSODBC)) > 0		&& CONTINUE WITH INTERFACE
	** UNIVERSAL DATA FOR BOTH INTERFACES
		STORE allt(OUR_MAN.invnum) TO lcO	&& gcORDER
		store allt(GCSYCD) to lcsycd		&& system id 
		store allt(OUR_MAN.uniqueid) to lcID&& uniqueid
		lcd1 = ""
		lcd2 = ""
		lcd3 = ""
		ldtotal = ""
		store substr(dtoc(OUR_MAN.date),7,4) to lcd1
		store substr(dtoc(OUR_MAN.date),1,2) to lcd2
		store substr(dtoc(OUR_MAN.date),4,2) to lcd3
		ldtotal = lcd1+lcd2+lcd3
		lcd = val(ldTotal)

	store SUBSTR(OUR_MAN.time,1,2) + SUBSTR(OUR_MAN.time,4,2) + padl(ALLT(STR(SEC(DATETIME()))),2,"0")  to lctime	&&shiptime

** processed before  this should only be for gcsycd 1, Trade 
		store "" to lcProc	&& default always for Trumble (gcsycd 2) must put logic for trade side as trade is swept several times per day
				
		if glProc=.t.  and gcSycd = "1" && if package was processed and voided in TRade, set flag to "V" so SAP knows it is an update.  
				store "V" to lcProc
		endif
		
		if glProc=.f. and gcSycd="1" &&new record, no void
			store " " to lcProc 
         endif
          lntotfrt = 0   
        STORE OUR_MAN.PAYMENT TO LNPAY
		STORE allt(OUR_MAN.packer) TO lcP			    && apss logon
		store allt(str(OUR_MAN.curbox)) to lcCB			&& current box TEXT
		store our_man.curbox to lnCB					&& current box NUMERIC
		STORE ALLT(OUR_MAN.shipper)+ALLT(OUR_MAN.upstyp) TO lcCarr	&& ship via code
		store allt(str(OUR_MAN.maxbox)) to lcMB			&& current box
		store allt(str(OUR_MAN.pieces)) to lcPc
		STORE allt(OUR_MAN.zip) TO lcZp					&& postalcode
		STORE OUR_MAN.wight TO lnW		&& WEIGHT
		store allt(OUR_MAN.zone) to lcZn
		store allt(OUR_MAN.picknum) to lcPk
		STORE allt(OUR_MAN.trknum) TO lcT				&& 
		
		STORE OUR_MAN.SHIPCHG TO lntotfrt &&this was the culprit that was sending the 6.00 fee twice to SAP
										
		STORE OUR_MAN.SHIPCHG  TO lnbase
										
		store allt(OUR_MAN.ucc128) to lcUcc

		store OUR_MAN.codchg to lncodchg		&&& amount of COD on package
		store OUR_MAN.CODfee to lnCod
		store OUR_MAN.dvi to lndvc	&&dvfee
		store OUR_MAN.dv to lndvia	&& dv amount
		store OUR_MAN.dvpip to lnaia	&& ALT /Pip dv amount
		store OUR_MAN.dvipip to lnaic	&& ALT/PIP INS FEE
		STORE OUR_MAN.AODFEE TO lnaodc	&& aod charges
		store OUR_MAN.tagfee to lnctc	&& call tag fee
		store OUR_MAN.hand to lnhf	&& handling fee

		IF allt(OUR_MAN.codf) = "Y"
			store "COD TYPE" + allt(str(OUR_MAN.anynum)) to lx	&& cod type for charges
		ELSE
			store "COD TYPE" to lx	&& non cod type for charges
		ENDIF

		STORE "" TO lcTmp1
		STORE "" TO lcTmp2
		STORE "" TO lcTmp3
		STORE "" TO lcTmp4
		STORE "" TO lcTmp5
		STORE "" TO lcTmp6
		STORE "" TO lcTmp7
		STORE "" TO lcTmp8
		STORE "" TO lcTmp9
		STORE "" TO lcTmp10
		STORE "" TO LCTMP11
		STORE "" TO lcTotal

*?lnbase
        iF LNPAY=1
		STORE "MNSEQ, MNMODE, MNPKG, MNLTL, MNWGT, MNZONE, MNCAR, MNTRCK, MNTXT1, " TO lcTmp2
		store "MNBASE, MNUCC, MNCOD, MNDVC, MNDVIA, MNAIA, MNAIC, MNAODC, MNCTC, " to lctmp3
		store "MNHFEE, MNNET, MNCODC, MNPOST) " to lctmp4

		STORE "VALUES (?lcd, ?lcTime, ?lcO, ?lcsycd, ?lcID, ?lcP, ?lcCB, ?lcCarr, ?lcMB, ?lcPc, " TO lcTmp5
		store "?lnW, ?lcZn, ?lcPk, ?lcT, ?lx, " to lctmp6
		STORE "?lnbase, ?lcUcc, ?lncod, ?lndvc, ?lndvia, ?lnaia, ?lnaic, ?lnaodc, ?lnctc, " to lctmp7
		store "?lnhf, ?lntotfrt, ?lncodchg, ?lcproc)" to lctmp8
        
         ELSE  &&NOT PREPAID-SEND ZERO'S
         
         STORE "MNSEQ, MNMODE, MNPKG, MNLTL, MNWGT, MNZONE, MNCAR, MNTRCK, MNTXT1, " TO lcTmp2
		store "MNBASE, MNUCC, MNCOD, MNDVC, MNDVIA, MNAIA, MNAIC, MNAODC, MNCTC, " to lctmp3
		store "MNHFEE, MNNET, MNCODC, MNPOST) " to lctmp4
		
	STORE "VALUES (?lcd, ?lcTime, ?lcO, ?lcsycd, ?lcID, ?lcP, ?lcCB, ?lcCarr, ?lcMB, ?lcPc, " TO lcTmp5
		store "?lnW, ?lcZn, ?lcPk, ?lcT, ?lx, " to lctmp6
		STORE "0.00, ?lcUcc, ?lncod, ?lndvc, ?lndvia, ?lnaia, ?lnaic, ?lnaodc, ?lnctc, " to lctmp7
		store "?lnhf, 0.00, ?lncodchg, ?lcproc)" to lctmp8	

    ENDIF
** END OF UNIVERSAL DATA
&& SAP ORDER
	**	IF ALLT(OUR_MAN.MSODBC)= "S" AND GCSYCD = "1"	&& TRADE
		
		IF GCSYCD == "1"	&& TRADE &&change because of intermittent problem
		*	wait window "trade" timeout .3&&TEST COMMENT OUT PRODUCTION 
			STORE "INSERT INTO PRD.prd.ZTRDAPSS(MNDTE, MNTIME, MNORD, MNSYCD, MNUNID, MNSYS, " TO lcTmp1
		ENDIF	
	**	IF ALLT(OUR_MAN.MSODBC)= "S" AND GCSYCD = "2"	&& TRUMBLE
        IF GCSYCD == "2"	&& TRUMBLE EXACTLY &&change because of intermittent problem
       *	 wait window "TRUMBLE" timeout .3 &&TEST COMMENT OUT PRODUCTION 
			STORE "INSERT INTO PRD.prd.ZAPSS(MNDTE, MNTIME, MNORD, MNSYCD, MNUNID, MNSYS, " TO lcTmp1
		ENDIF	

		STORE lcTmp1+lcTmp2+lcTmp3+lcTmp4+lcTmp5+lcTmp6+lcTmp7+lcTmp8 TO lcTotal

		gnsqlerror2 = 0
** INSERT SQL STATEMENT
		IF ALLT(OUR_MAN.MSODBC)= "S" AND GNSQLSAP >0
			gnsqlerror=SQLEXEC(gnsqlSAP,lcTotal)
		ENDIF
*** END OF INSERT STATEMENT
		IF gnSqlError>0
			wait window "Saving record data.."+ALLTRIM(gcOrder) TIMEOUT .3  nowait
            		ELSE
			MESSAGEBOX("Bad insert result code for 'header' record..."+str(gnSqlError))
		select(lnoldsele)
		RETURN .F.
 		ENDIF
*
		IF GNSQLSAP < 1
			messagebox("SAP ODBC connection not available."+chr(13)+"You must EXIT SPSS and Re-enter to reconnect." +;
			chr(13)+"Any shipments processed WILL NOT transfer information to SAP.", 0,"     ODBC CONNECTION NOT AVAILABLE")
		ENDIF
ELSE
   
	* MESSAGEBOX("MANUAL SHIPMENT, DATA NOT RETURNED TO INTERFACE")
ENDIF	&& interfaced shipment, OUR_MAN.MSODBC O FOR AS400, S FOR SAP

ENDSCAN

ENDIF &&not a future ship

*!*	MESSAGEBOX("REMINDER-TEST SYSTEM",0 + 64,"NAG MESSAGE")
SELECT(LNOLDSELE)

return

***************************************************************************
function blanktxt()
***************************************************************************
*** FOR RR DONNELLY FORM GENERATOR LABEL
** WONT PRINT UNLESS ANYTEXT2 FIELD HAS DATA IN IT.

If LEN(ALLT(man.anytext2))> 0
lcReturnx = "^FT25,1250^A0N,50,45^FDSent to you " + allt(man.anytext2)+"^FS"	
ELSE
lcReturnx = ""
endif

return lcReturnx


***************************************************************************
FUNCTION SAPDATA()
***************************************************************************
**updated with AFFES, BBB, GANDER, Marshall Fields

	PUBLIC glucc  &&determines gander mtn order
	glucc =.f.
	gcOrder = padl(allt(gcOrder),10,"0")&&adds zero 808123456=0808123456
	lcstore=" "

	replace batch.pkgid with ALLT(gcOrder)+ ".01"	&& b_batchmode 0808123456.01
	replace batch.invnum with ALLT(gcOrder)	&& b_batchmode  0808123456
	replace batch.dept with ALLT(gcOrder)	&& b_batchmode  0808123456
		
*** get data from VBUK ( 'Sales document status data' )
	gnSqlError=0
	gnSqlError=SQLEXEC(gnSqlsap,"SELECT PRD.prd.VBUK.WBSTK, PRD.prd.VBUK.MANDT FROM PRD.prd.VBUK WHERE PRD.prd.VBUK.VBELN =?gcOrder and PRD.prd.VBUK.MANDT = '400'")

	IF gnSqlError<0
	*brow
   *  WAIT WINDOW "FIRST ERROR CHECK" 	
		WAIT window "gnSqlError...VBUK "+str(gnSqlError) TIMEOUT 3
	ENDIF

	IF USED('sqlresult') AND RECCOUNT('sqlresult')>0
		STORE SELECT() to lnOldSele
			lcG_Status=""
				STORE SQLRESULT.wbstk TO lcG_Status	&& 'Goods issue status'
					PUBLIC gcMandt
						gcMandt=""
							STORE SQLRESULT.mandt TO gcMandt	&& 'Client'
					IF USED('sqlresult')
						USE IN SQLRESULT
					ENDIF
							SELECT (lnOldSele)
				ELSE
				*BROW
				*WAIT WINDOW "2ND ERROR CHECK"
					MESSAGEBOX ("Delivery number not found in VBUK")
				RETURN .F.
	ENDIF
*	end of VBUK table
*** get data from LIKP ( 'Delivery header data' )
	gnSqlError=0
	gnSqlError=SQLEXEC(gnSqlsap,"SELECT PRD.prd.LIKP.BTGEW, PRD.prd.LIKP.GEWEI, PRD.prd.LIKP.KUNNR, PRD.prd.LIKP.MANDT, PRD.prd.LIKP.INCO1, PRD.prd.LIKP.KUNAG FROM PRD.prd.LIKP WHERE PRD.prd.LIKP.VBELN=?gcOrder and PRD.prd.LIKP.MANDT = '400'")

	IF gnSqlError<0
		WAIT window "gnSqlError...LIKP "+str(gnSqlError) TIMEOUT .5
	ENDIF

	IF USED('sqlresult') AND RECCOUNT('sqlresult')>0

		STORE SELECT() to lnOldSele
		SELE SQLRESULT
		lcKUNNR = ""
		STORE SQLRESULT.KUNNR to lcKUNNR	&& customer #


		IF GCSYCD = '1' && TRADE

			ELSE
				replace batch.scntrcode with SQLRESULT.INCO1		&& b_batchmode
		ENDIF

		PUBLIC gcSoldTo
		STORE SQLRESULT.kunag TO gcSoldTo	&& 'Sold-to party'; used in pack app preferences

		IF USED('sqlresult')
			USE IN SQLRESULT
		ENDIF
		SELECT (lnOldSele)
	ELSE
		MESSAGEBOX ("Delivery number not found in LIKP")
		RETURN .F.
	ENDIF
*	end of LIKP table

*** get data from VBPA ( 'Sales document partner data' )
	gnSqlError=0

	lctmp1 = "SELECT PRD.prd.VBPA.MANDT, PRD.prd.VBPA.POSNR, PRD.prd.VBPA.PARVW, PRD.prd.VBPA.ADRNR, PRD.prd.VBPA.KUNNR " 
	lctmp2 = "FROM PRD.prd.VBPA WHERE PRD.prd.VBPA.VBELN=?gcOrder AND PRD.prd.VBPA.POSNR=0 and PRD.prd.VBPA.PARVW='WE'and PRD.prd.VBPA.MANDT = '400'"

	lcTOTAL = lctmp1+lctmp2

	gnSqlError=SQLEXEC(gnSqlsap,lcTOTAL)

	IF gnSqlError<0
		WAIT window "gnSqlError...VBPA "+str(gnSqlError) TIMEOUT .5
	ENDIF

	IF USED('sqlresult') AND RECCOUNT('sqlresult')>0
		STORE SELECT() to lnOldSele
		SELE SQLRESULT

		lcAddNum=""	&& combine with street name later
		STORE SQLRESULT.adrnr TO lcAddNum	&& address number
		
        REPLACE BATCH.BILLTO WITH SQLRESULT.KUNNR
        
          if batch.billto="0001097587"   &&&Marshall Fields
          replace batch.cscode with sqlresult.kunnr
          endif

		IF USED('sqlresult')
			USE IN SQLRESULT
		ENDIF
		SELECT (lnOldSele)
	ELSE
		MESSAGEBOX ("Delivery number not found in VBPA")
		RETURN .F.
	ENDIF
*	end of VBPA table

*  GET DATA FROM LIPS
	gnSqlError = SQLEXEC(gnSqlsap,"select PRD.prd.LIPS.VGBEL from PRD.prd.LIPS where PRD.prd.LIPS.MANDT = '400' and PRD.prd.LIPS.VBELN = ?gcorder")
	STORE SQLRESULT.vgbel to gcSalesOrd
** END OF LIPS
***********************
**new stuff for affes &&&has to be after LIPS TO GET THE SALES ORDER NUMBER AND THE AFFES # IS ACTUALLY 0000001259, NOT 101259!!
	gnSqlError=0

	gnSqlError=SQLEXEC(gnSqlsap,"SELECT PRD.prd.VBPA.KUNNR FROM PRD.prd.VBPA WHERE PRD.prd.VBPA.VBELN=?gcSalesOrd AND PRD.prd.VBPA.POSNR=0 and PRD.prd.VBPA.PARVW='RE'and PRD.prd.VBPA.MANDT = '400'")

	IF gnSqlError<0
		WAIT window "gnSqlError...NEW STUFF "+str(gnSqlError) TIMEOUT .5
	ENDIF
*BROW
	IF USED('sqlresult') AND RECCOUNT('sqlresult')>0

		STORE SELECT() to lnOldSele
		SELE SQLRESULT
		lcNEWKUNNR = ""
		STORE SQLRESULT.KUNNR to lcNEWKUNNR	&& customer #
   
     IF LCNEWKUNNR="0000001259"  && emailed John for confirmation-THIS IS CORRECT FOR THE 'RE' IS THE BILL TO
*!*	       			replace batch.shipper WITH "FEXP"  &&CHANGED 02/11/09 per Tom
*!*	      			REPLACE BATCH.UPSTYP WITH "G"
*!*	      			REPLACE BATCH.PAYMENT WITH 4
      			gcAFFES="Y"
      		ENDIF
      		gcAffes="N"
     ENDIF
  ***next, need store # if it's an AFFES order
  IF gcAFFES="Y"
  	**GET THE KEY FROM VBPA TO GO TO KNB1
  
  gnSqlError=0

  gnSqlError=SQLEXEC(gnSqlsap,"SELECT PRD.prd.VBPA.KUNNR FROM PRD.prd.VBPA WHERE PRD.prd.VBPA.VBELN=?gcSalesOrd AND PRD.prd.VBPA.POSNR=0 and PRD.prd.VBPA.PARVW='AG'and PRD.prd.VBPA.MANDT = '400'")

		IF gnSqlError<0
			*WAIT window "gnSqlError...AFFES KEY "+str(gnSqlError) TIMEOUT .5
		ENDIF
*BROW
		IF USED('sqlresult') AND RECCOUNT('sqlresult')>0

		STORE SELECT() to lnOldSele
		SELE SQLRESULT
  
		lcAfkey = ""
		STORE sqlresult.kunnr TO lcAfkey  
  			ELSE
  			*WAIT WINDOW "NO KEY FOR AFFES" TIMEOUT 2
  		ENDIF
    **GO TO KNB1 WITH KEY
    
     gnSqlError=0

 	 gnSqlError=SQLEXEC(gnSqlsap,"SELECT PRD.prd.KNB1.EIKTO FROM PRD.prd.KNB1 WHERE PRD.prd.KNB1.KUNNR=?lcAfkey AND PRD.prd.KNB1.BUKRS='0010'and PRD.prd.KNB1.MANDT = '400'") &&THIS IS WHERE I NEED TO FINISH-AG?? 

		IF gnSqlError<0
			WAIT window "gnSqlError...KNB1 WITH AFFES KEY "+str(gnSqlError) TIMEOUT .5
		ENDIF
*BROW
		IF USED('sqlresult') AND RECCOUNT('sqlresult')>0

		STORE SELECT() to lnOldSele
		SELE SQLRESULT
    
		lcAfStore = ""    
    
    	STORE sqlresult.eikto  TO lcAfStore
    	REPLACE BATCH.CSNAM2 WITH ALLTRIM(lcAfStore)   && figure out which field to store it to 12 char
    	*	WAIT WINDOW "STORE NUMBER=   "+ALLTRIM(LCaFSTORE) 
    		ELSE
    		WAIT WINDOW "NO STORE NUMBER FOR AFFES"
    		
    	ENDIF	
     ELSE &&not AFFES
     
  ENDIF 
          
**********END OF JUST AFFES	STUFF
**new for GANDER, BBB AND AFFES TO TRIGGER UCC128 #'S FROM CUSTOMER DATABASE
	gnSqlError=0

	gnSqlError=SQLEXEC(gnSqlsap,"SELECT PRD.prd.VBPA.KUNNR FROM PRD.prd.VBPA WHERE PRD.prd.VBPA.VBELN=?gcSalesOrd AND PRD.prd.VBPA.POSNR=0 and PRD.prd.VBPA.PARVW='RG'and PRD.prd.VBPA.MANDT = '400'")
IF gnSqlError<0
		WAIT window "gnSqlError...NEW GANDER STUFF "+str(gnSqlError) TIMEOUT .5
	ENDIF

	IF USED('sqlresult') AND RECCOUNT('sqlresult')>0

		STORE SELECT() to lnOldSele
		SELE SQLRESULT
		lcNEWKUNNR = ""
		STORE SQLRESULT.KUNNR to lcNEWKUNNR	&& customer #
		
	IF LCNEWKUNNR="0000113515"  &&Gander Mtn

       		replace batch.billto WITH "0113515"
       		replace batch.cscode WITH "0113515"
       		gcGander="Y"
       					ELSE
   		ENDIF
      	
   	IF LCNEWKUNNR="0002003652" &&Bed Bath & Beyond

       		replace batch.billto WITH "0002003652"
       		replace batch.cscode WITH "0002003652"
       					ELSE
       	ENDIF		
   
   	IF LCNEWKUNNR="0000101259"  &&AFFES
   	
       		replace batch.billto WITH "0000101259"
       		replace batch.cscode WITH "0000101259"
       							ELSE
    ENDIF	
     ENDIF	
*** get data from VBAK ( 'Sales document header data' )
	gnSqlError=0
	gnSqlError=SQLEXEC(gnSqlsap,"SELECT PRD.prd.VBAK.NETWR, PRD.prd.VBAK.BNAME,PRD.prd.VBAK.AUART, PRD.prd.VBAK.BSTNK, PRD.prd.VBAK.VKBUR FROM PRD.prd.VBAK WHERE PRD.prd.VBAK.MANDT = '400' and PRD.prd.VBAK.VBELN=?gcSalesOrd")

	IF gnSqlError<0
		WAIT window "gnSqlError...VBAK "+str(gnSqlError) TIMEOUT 5
	ENDIF

	IF USED('sqlresult') AND RECCOUNT('sqlresult')>0
		STORE SELECT() to lnOldSele
		SELE SQLRESULT

		*replace	batch.ponum with SUBSTR(SQLRESULT.bstnk,1,8)	&&&ORIG LP 10/26/07 b_batchmode PO number
		replace	batch.ponum with ALLTR(SQLRESULT.bstnk) &&
	
* CODAMT
		lncodamt = 0
		STORE SQLRESULT.NETWR TO lncodamt
		STORE SQLRESULT.bname to gccust_ref			&& cust_ref for wallmart
		lcx = ""
		STORE allt(upper(SQLRESULT.auart)) to lcx	&& looking for order type to check for vlaidation

		IF USED('sqlresult')
			USE IN SQLRESULT
		ENDIF

		SELECT (lnOldSele)

	ELSE
		MESSAGEBOX ("Delivery number not found in VBAK")
	ENDIF	&&	end of VBAK table

** GET DATA FROM VBKD	
** see if it is a cod OR DEPT FOR GANDER MTN SHIPMENT
LCPODEPT = ""
gnSqlError=SQLEXEC(gnSqlsap,"SELECT PRD.prd.VBKD.ZTERM, PRD.prd.VBKD.BSTKD FROM PRD.prd.VBKD WHERE PRD.prd.VBKD.MANDT = '400' and PRD.prd.VBKD.VBELN=?gcSalesOrd")	
	IF gnSqlError<0
		WAIT window "gnSqlError...VBKD "+str(gnSqlError) TIMEOUT 5
	ENDIF
	SELE SQLRESULT
	STORE ALLT(SQLRESULT.ZTERM) TO LCCODFLAG
	IF !EMPTY(SQLRESULT.BSTKD) AND gcGander="Y"
			STORE ALLTRIM(SQLRESULT.BSTKD) TO LCPODEPT &&podept
					**this is where it needs to be parsed out
		lnBegin = 0
       	lcPo = ""
       	lcDept = ""
       	 **parse out PO 
       	STORE lcPodept TO gcString
       	STORE "-" TO gcFindstring
       	lnBegin = ATC(gcFindstring, gcString)
       	lcPo = SUBSTR(LEFT(lcPodept, lnBegin-1),1) 
       
       replace batch.ponum WITH ALLTRIM(lcPo)
        
       **now get dept
       
       lcDept = SUBSTR(lcPodept,lnbegin+1,2)
       **lcDept=right(lcPodept,2)
                
       REPLACE BATCH.anytext WITH ALLTRIM(lcDept)
         glUcc = .t.
       
	ENDIF
	    ***  COD FLAG

	DO CASE	
		CASE LCCODFLAG = "08"
           * WAIT WINDOW "08 ZTERM = " + LCCODFLAG
			replace batch.anynum with 0				&&& b_batchmode COD TYPE STORED IN ANYNUM FLAG FOR RETURN TO SAP TABLE
			replace batch.codamt with lncodamt		&&& b_batchmode 
*						
		CASE LCCODFLAG = "12"
			*WAIT WINDOW "12 ZTERM = " + LCCODFLAG
			replace batch.anynum with 2				&&& b_batchmode COD TYPE STORED IN ANYNUM FLAG FOR RETURN TO SAP TABLE
			replace batch.codamt with lncodamt		&&& b_batchmode TOTAL COD AMNT

		CASE LCCODFLAG = "58"      &&Christmas COD
			*WAIT WINDOW "58 ZTERM = " + LCCODFLAG
			replace batch.anynum with 0				&&& b_batchmode COD TYPE STORED IN ANYNUM FLAG FOR RETURN TO SAP TABLE
			replace batch.codamt with lncodamt		&&& b_batchmode TOTAL COD AMNT

	ENDCASE
**END OF COD FLAG SETTING
*** get data from ADRC ( 'Central address management' )
*
			gnSqlError=0

			lctmp1 = "SELECT PRD.prd.ADRC.CLIENT, PRD.prd.ADRC.NAME1, PRD.prd.ADRC.NAME2, PRD.prd.ADRC.NAME3,PRD.prd.ADRC.NAME4, PRD.prd.ADRC.TITLE, "
			lctmp2 = "PRD.prd.ADRC.STREET,PRD.prd.ADRC.CITY1, PRD.prd.ADRC.REGION, PRD.prd.ADRC.POST_CODE1, PRD.prd.ADRC.TEL_NUMBER, PRD.prd.ADRC.COUNTRY, PRD.prd.ADRC.STR_SUPPL1, "
			lctmp3 = "PRD.prd.ADRC.STR_SUPPL2 FROM PRD.prd.ADRC WHERE PRD.prd.ADRC.CLIENT='400' and PRD.prd.ADRC.ADDRNUMBER = ?lcaddnum"
			lctmp4 = ""
			lctmp5 = ""

			lcTOTAL = lctmp1+lctmp2+lctmp3+lctmp4+lctmp5
			gnSqlError=SQLEXEC(gnSqlsap,lcTOTAL)

			IF gnSqlError<0
				WAIT window "gnSqlError...ADRC "+str(gnSqlError) TIMEOUT 1
			ENDIF

			IF USED('sqlresult') AND RECCOUNT('sqlresult')>0

				STORE SELECT() to lnOldSele
				SELE SQLRESULT
*** populate screen ( or other ) variables

			IF GCSYCD = "2" 	&& TRUMBLE
		
				if len(allt(sqlresult.name3)) >0
					replace batch.company with allt(SQLRESULT.name3)	&&B_BATCHMODE Consignee company name

					if len(allt(title)) >0	&& is title present?
					replace batch.csnam2 with allt(title) + " " + allt(SQLRESULT.name2) + " " + allt(SQLRESULT.name1)	&&B_BATCHMODE 
					else	&& 
						replace batch.csnam2 with allt(SQLRESULT.name2) + " " + allt(SQLRESULT.name1) &&Consignee company name
					endif
				else
					if len(allt(title)) >0	&& is title present?t
						replace batch.company with allt(title) + " " + allt(SQLRESULT.name2) + " " + allt(SQLRESULT.name1)  && b_batchmode Consignee contact
					else	&& 
						replace batch.company with allt(SQLRESULT.name2) + " " + allt(SQLRESULT.name1)	&&B_BATCHMODE
					    REPLACE BATCH.CSNAM2 WITH "ATTN"
					endif
				ENDIF
								                  
				replace batch.address with allt(SQLRESULT.street)		&&B_BATCHMODE ADDRESS
				replace batch.csadd2 with allt(SQLRESULT.STR_SUPPL1)	&&B_BATCHMODE
				replace batch.csadd3 with allt(SQLRESULT.STR_SUPPL2)	&&B_BATCHMODE
				replace batch.cscity with allt(SQLRESULT.city1)	&&B_BATCHMODE Consignee city
				replace batch.csstat with allt(SQLRESULT.region)	&&B_BATCHMODE STATE
				replace batch.zip with allt(SQLRESULT.post_code1)	&&B_BATCHMODE ZIP CODE
				replace batch.cntcode with allt(SQLRESULT.country)	&&B_BATCHMODE COUNTRY
				*WAIT WINDOW "COUNTRY =   "+ALLTRIM(BATCH.CNTCODE) TIMEOUT 2
				if allt(SQLRESULT.country) == "C"
					replace batch.cntcode with "CA"	&& canadian shipment		
				endif 
			ENDIF

			IF GCSYCD = "1" 	&& TRADE
			** COMPANY IS IN DIFFERENT FIELDS FOR TRADE
			** CUS_2 IS REQUIRED 
				if len(allt(sqlresult.name1)) >0  && DOES COMPANY EXIST?
					replace batch.company with SQLRESULT.name1	&&B_BATCHMODE
			                                    *   BBB ORDER-GET STORE NUMBER FROM COMPANY
			 IF LCNEWKUNNR="0002003652" &&BBB
							store alltrim(batch.company) to lcstore
							replace batch.csnam2 with substr(lcstore,16,8)
							ENDIF
				
						IF LCNEWKUNNR="0000113515" &&GANDER
							store alltrim(batch.company) to lcstore
							replace batch.csnam2 with substr(lcstore,16,8)
							ENDIF
							
				else
	replace batch.company with allt(title) + " " + allt(SQLRESULT.name2) + " " + allt(SQLRESULT.name1)	&&B_BATCHMODE
				endif
			
			    if s_shipper<>"POM" or "POS"  
				replace batch.address with allt(SQLRESULT.street)		&&B_BATCHMODE STREET
				replace batch.csadd2 with allt(SQLRESULT.STR_SUPPL1)	&&B_BATCHMODE  ADDRESS2
				   if batch.csadd2="P.O."             &&GET RID OF SOME OF THE PO BOX NUMBERS 
				     replace batch.csadd2 with ""
				     endif
				   if batch.csadd2="PO"  
				     replace batch.csadd2 with ""
				     endif
                   if batch.csadd2="P O" 
				     replace batch.csadd2 with ""
			     endif
			
				replace batch.csadd3 with allt(SQLRESULT.STR_SUPPL2)	&&B_BATCHMODE ADDRESS 3
				Replace batch.cscity with allt(SQLRESULT.city1)	&&B_BATCHMODE
				replace batch.csstat with allt(SQLRESULT.region)	&&B_BATCHMODE CONSIGNEE STATE
				replace batch.zip with allt(SQLRESULT.post_code1)	&&B_BATCHMODE CONSIGNEE ZIP
				endif
				
				replace batch.cntcode with allt(SQLRESULT.country)	&&B_BATCHMODE Consignee city
				*WAIT WINDOW "COUNTRY =   "+ALLTRIM(BATCH.CNTCODE) TIMEOUT 2
				if allt(SQLRESULT.country) == "C"
					replace batch.cntcode with "CA"	&& canadian shipment		
				endif 
				replace batch.phone with allt(SQLRESULT.tel_number)	&&B_BATCHMODE
             endif
*
				IF USED('sqlresult')
					USE IN SQLRESULT
				ENDIF
				SELECT (lnOldSele)

			ELSE
				MESSAGEBOX ("Delivery number not found in ADRC")
				RETURN .F.
			ENDIF
*	end of ADRC table
*** END OF SAPDATA
return


****************************************************************************************************************
FUNCTION pkgid_lookup
****************************************************************************************************************
PARAMETERS lcPkgId

lnOldSele = SELECT()

*WAIT WINDOW "PKGID LOOKUP .." + LCPKGID TIMEOUT 1
llOK=.T.
lcCarrier=""

lnRecs=ADIR(laMan,'???man*.dbf')	&& get names of all files with MAN in them
***puts them in laMan array, ln rec gets # of *man* files.

FOR I = 1 TO lnRecs
*	WAIT WINDOW "Searching... "+ALLT(laMan(1,1)) NOWAIT
		
	lcFileName=laMan(I,1)
	USE(lcFileName) IN 0 AGAIN SHARED ALIAS MS_MAN
	
	SELECT INVNUM, shipper, shpsts FROM ms_man WHERE ALLT(ms_man.INVNUM)= ALLT(lcPkgId) INTO CURSOR cShpSts
			
	SELE cShpSts
*	WAIT WINDOW " INVOICE NUM  " + CSHPSTS.INVNUM + "CARRIER " + CSHPSTS.SHIPPER TIMEOUT .5
*	BROW

		IF reccount('cShpSts')>0			
			glProc = .t.  && this has been processed before must set update flag for sap insert

			SCAN		&& look through all records 
				IF ALLT(cShpSts.shpsts)<>"V"	&&&&& if not voided
					llOK=.F.		&& found it exit

				*	WAIT WINDOW "CARRIER IS .."+ ALLT(cShpSts.shipper)

					EXIT		&& exit scan if we found a bad one, jumps to PAST NEXT 
				ENDIF
			ENDSCAN
	
			lcCarrier=ALLT(cShpSts.shipper)
			
			I=lnRecs	&& stop processing since we found the correct MAN file
			
		ENDIF

	IF USED('cShpSts')
		USE IN cShpSts
	ENDIF		

	IF USED('ms_man')
		USE IN ms_man
	ENDIF
	
NEXT I

WAIT CLEAR

SELECT (lnOldSele)

RETURN llOK
********************************************************************************************
FUNCTION pkgid_lookup2		&& modified 9/15/05
*********************************************************************************************
PARAMETERS lcPkgId

lntmpsel = SELECT()
*WAIT WINDOW "PKGID LOOKUP .." + LCPKGID TIMEOUT 1
llOK=.T.
lcCarrier=""
lcfilename = ""


*IF ! USED('manfiles')
IF ! USED('MS_SHIPPER')  &&CHANGED FOR ERROR OF FILE IN USE
	USE MS_shipper SHARED IN 0 ALIAS our_shipper
ENDIF	
	SELECT our_shipper
	SELECT DISTINCT manfil FROM our_shipper WHERE active = "Y" INTO CURSOR manfiles
	
	IF USED('our_shipper')
		USE IN our_shipper
	endif
*ENDIF

SELECT manfiles
GO top
SCAN 
	lcfilename = ALLTRIM(manfiles.manfil)
	USE(lcFileName) IN 0 AGAIN SHARED ALIAS MS_MAN

	SELECT INVNUM, shipper, shpsts FROM ms_man WHERE ALLT(ms_man.INVNUM)= ALLT(lcPkgId) INTO CURSOR cShpSts
			
	SELE cShpSts
	*brow
		IF reccount('cShpSts')>0			
			glProc = .t.  && this has been processed before must set update flag for sap insert
			SCAN		&& look through all records 
				IF ALLT(cShpSts.shpsts)<>"V"	&&&&& if not voided
					llOK=.F.		&& found it -exit
				*	WAIT WINDOW "CARRIER IS .."+ ALLT(cShpSts.shipper)
				 MESSAGEBOX("ALREADY PROCESSED, THE SAME INVOICE NUMBER FOUND IN CARRIER "  + CHR(13)+ ALLT(CSHPSTS.SHIPPER),0+48,"WARNING-DUPLICATE!!!")
					EXIT		&& exit scan if we found a bad one, jumps to PAST NEXT 
				ENDIF
			ENDSCAN
	
			lcCarrier=ALLT(cShpSts.shipper)
		ENDIF
* cleanup
	IF USED('cShpSts')
		USE IN cShpSts
	ENDIF		

	IF USED('ms_man')
		USE IN ms_man
	ENDIF
endscan	

SELECT (lntmpsel)

RETURN llOK

*****************************************************************************
function msucc() 
*****************************************************************************
** create customer # called BEDBATH and give them a starting UCC128# of 
lnold= select()

if empty(man.ucc128)
	if ! used('OUR_CUST')
		use customer in 0 again shared alias our_cust
	endif

	sele our_cust
	locate for allt(our_cust.csnum) == "BEDBATH"

	IF FOUND()
		gcMSUCC = our_cust.csucc128
		REPLACE OUR_CUST.CSUCC128 WITH MOD10(OUR_CUST.CSUCC128,20,.T.,"W")
		sele man
		replace man.ucc128 with gcMSUCC
		replace man.anytext with gcmsucc &&lp added 11/16/04 for BBB requirements
	else
		messagebox("CUSTOMER UCC IS NOT AVAILABLE")
	ENDIF
	*messagebox("custnum = " + GcCUSTNUM + " ucc# =  " + gcmsucc)
	if used('OUR_CUST')
		use in our_cust
	endif
ELSE
	gcMSUCC = allt(man.ucc128)
endif

select(lnold)
RETURN "x"
****************************************************************
FUNCTION B_SHIP() && set anytext2 for dimchart
****************************************************************
sa_cust(1)=""
IF GCSYCD="1"
STORE "N" TO SC_SHIP.A_SHIP(63)
ENDIF
***************************************************************
Function b_shiptrum()
*****************************
sa_cust(3) = 0


******************************************************************
function B_batch()    &&NEW for filelink interface  6/30/04
******************************************************************

IF GCSYCD = "1"	&& void not available in trade shipping screens
	sf_ship.p_void.enabled = .f.
endif

glOKtoship = .t.

public gnLoop
gnLoop = 0
glProc = .f.	&& has this been processed yet

replace batch.ckcash with "K"
REPLACE BATCH.ANYTEXT2 WITH "N" &&DO THIS IN B_SHIP INSTEAD?
 
sa_cust(5) = ""		&& ODBC FLAG
sa_cust(8) = ""		&& MAN.uccstornum  STORE #
sa_cust(10)= ""



STORE SELECT() to lnOldSele
*** call GetValTxt to get Delivery Number ***
*
If  gnsqlsap > 0 && make sure SAP conn are available

	gcOrder = ""
	gcOrder=GETVALTXT("ORDER NUMBER ","Enter Order Number: ", "",14,,"",,"K")


	IF EMPTY(ALLT(gcOrder))
		RETURN .T.
	ENDIF


	gcorder = allt(gcorder)
	if len(gcorder) = 8 and substr(gcorder,1,1) = "8"		
	
	*** CHECK FOR SHIPMENT PROCESSING
		llok = .t.
	** call pkgid_lookup function to see if order is already processed and not voided.
	** changed PJ 9/15/05 to pkgid_lookup2

		pkgid_lookup2(padL(gcOrder,10,"0")) &&THIS IS WHERE WE LOSE IT 06/22/04
	
		IF ! llok	&&  false
			messagebox("ALL PACKAGES WITH THIS ORDER MUST BE VOIDED BEFORE PROCEEDING" ;
			+CHR(13) + "PLEASE VOID ALL PACKAGES AND TRY AGAIN", 0+16, "MUST VOID ORDER")
	
			return .f.	&& CANCELS SHIPMENT
		endif
 		&& CALL sap FUNCTION TO GET sap DATA. then exits this function. 
		sapdata()		
		sa_cust(5)= 'S'		&&Calls
		return .t.
	endif
******otherwise, it is a Trumble shipment		
		gcOrder=padr(gcOrder,14," ")
	*
		replace batch.pkgid with ALLT(gcOrder)+ ".01"		&&b_batchmode
		replace batch.invnum with gcOrder					&&b_batchmode
		replace batch.dept with allt(gcOrder)			&& b_batchmode PKGID

	else

		if gnsqlSAP <1		
			messagebox("SAP ODBC connection not available."+chr(13)+"You must EXIT SPSS and Re-enter to reconnect." +;
			chr(13)+"You will be unable to retrieve any information from SAP.", 0,"   SAP ODBC CONNECTION NOT AVAILABLE")
		endif
				
endif	&&SAP connection available

*
SELECT(lnOldSele)

KEYBOARD "{ENTER}"

RETURN  &&end of B_Batch()

************************************************************************************
function a_batch
************************************************************************************

IF ! glOKtoship
	LNANSWER = MESSAGEBOX("Order not found."+CHR(13) + "Try Again?", 4+32 , "ORDER NOT FOUND!" ) && YESNO DEFLT YES

	IF LNANSWER = 6 && YES
		SC_SCAN.STAYINLOOP = .T.
		RETURN .F.
	ELSE && NO
		RETURN .F.
	ENDIF
ENDIF

IF gcSycd = "1"
	replace batch.comsresd with "Commercial"
		ELSE
	replace batch.comsresd WITH "Residential"
ENDIF		

*brow


IF S_SHIPPER="FEXP"         &&WAS NOT CHANGING IT FOR fexp IN B_BATCH
 STORE "K" TO SC_SHIP.A_SHIP(65)
 ENDIF

RETURN

************************************************************************************
function B_PACK
************************************************************************************
 lncod=0 &&&fix for cod 'carry over'?
if s_shipper="POM"
 SC_SHIP.A_CUS(8)="80301"
 ENDIF
 
 ************************************************************************************
 FUNCTION A_SPECSERV
 ************************************************************************************
*WAIT WINDOW "BEGINING AFTER SPECIAL SERVICES" TIMEOUT 1
 * SC_SHIP.A_SHIP(65)="K"
  
  lcdept=""
  IF ALLTR(substr(l_pkid,1,10))<>ALLTR(SUBSTR(sc_ship.a_ship(84),1,10))&&department

 Messagebox("NUMBERS DON'T MATCH", 0+16, "SHIPMENT ERROR")
	Messagebox("Invoice Number MUST Match Department-CANCEL & RE-SCAN PICK SLIP!!", 16, "SHIPMENT MUST BE SCANNED AGAIN")
			keyboard '{ESC}'
			keyboard '{Y}'
				
		endif

**********************************************************************************
Function DimImport ()
**********************************************************************************
***Start with Importing the Excel Worksheet, saving as Dimchart.dbf
*  Function DimImport() brings in data from the xls, stored in Dimchart.dbf
*  Function DimChart looks up the dimensions &  calcualtes OS1 & OS2 Based
*  on dimensions.

Parameter lcTemp

lnOldSele= Select() &&keep track of current workarea

If File('Dimchart.xls')  && does user file exist?

    If SYS(2001,'safety')='ON'
        DELE FILE DIMCHART.DBF
           ELSE
        SET SAFETY OFF
        DELE FILE DIMCHART.DBF
        SET SAFETY ON
    ENDIF       
 ENDIF
 
** WAIT WINDOW "IMPORTING DIMCHART DATA..." TIMEOUT 1
 
 SELE 0
 IMPORT FROM DIMCHART.XLS TYPE XLS &&CREATE dIMCHART.DBF FROM DIMCHART.XLS
 *wAIT wINDOW "IMPORT COMPLETE"  tIMEOUT 1
 
 iF USED ('DIMCHART') &&CLOSE FILE-WILL OPEN IT LATER WHEN NEEDED
    USE IN DIMCHART
 ENDIF

SELECT (LNoLDsELE)  &&RE-ESTABLISH WORKAREA

RETURN ""

ENDfUNC
*******************************************************************************
FUNCTION DimChart () 
*******************************************************************************
***Assigned it to after rating once per package, add it
***in field level also?
***placed as a VFP COMMAND DIMCHART() -AFTER RATING, ASSIGN AT FIELD LEVEL ALSO??
***ADD IN USERTEXT 2 TO SCREEN, ADD OS PKG & DISPLAY SIZE BOX TO VIEW MANIFEST SCREEN, 
**MAKE SURE THAT KEYBOARD BUILDER IS ASSIGNED TO THE DIMWT KEY

lcusrtxt2=alltr(sc_ship.a_ship[63])

IF FILE('DimChart.Dbf')	&& does the file exist?
*
** found the file; open it and do the lookup
*

	lnOldSele = SELECT()

	USE DimChart IN 0 shared
	SELE DimChart
	
	lnweight=SC_SHIP.A_SHIP[8]

	GO TOP
   Locate for ALLTR(A)==ALLTR(LCUSRTXT2)    && TRY ALLTR INSTEAD OF VAL?
	
	IF ! FOUND()	&& couldn't find a match  
		 MESSAGEBOX ("CODE MUST BE ENTERED IN CORRECTLY, PLEASE CHECK CODE OR ENTER DIMENSIONS MANUALLY!! ",+0+16, "CODE WAS NOT FOUND")&& OR ENTERED INCORRECT, 
	 	 KEYBOARD'~'
	EndIf
		
	IF FOUND()
	*WAIT WINDOW "FOUND" TIMEOUT .5
	
   		 store INT(VAL(F)) to SC_SHIP.A_DIMWGT[1]         &&length sc_ship.a_dimwgt[1]
       	 store INT(VAL(G)) to sc_ship.a_dimwgt[2]         &&width sc_ship.a_dimwgt[2]
      	 store INT(VAL(H)) to  SC_SHIP.A_DIMWGT[3]        &&height SC_SHIP.A_DIMWGT[3]
         STORE INT(VAL(D)) TO  SC_SHIP.A_ship[9]          &&weight SC_SHIP.A_SHIP[9]
	
	  KEYBOARD 'W'
	  KEYBOARD '{ENTER}'
	 
	ENDIF
	    
     IF USED('DimChart')
	USE IN DimChart
		ENDIF

		RETURN 	&& no need to continue

	ENDIF
ENDFUNC

**		
****************************************************************************
FUNCTION COPYSHIPPER()
****************************************************************************		
	lnOldSele = SELECT()
IF FILE('SHIPPER.Dbf')	&& does the file exist?
*
*
IF !USED('shipper')
 USE shipper SHARED 
 ENDIF 
	SELECT SHIPPER	
	COPY TO MS_SHIPPER.DBF
	
	IF USED ('SHIPPER')	
	USE IN SHIPPER
	ENDIF
	
ENDIF	
SELECT (lnoldsele)
	RETURN
	
**********************************************************************
FUNCTION b_packtrade
***********************************************************************
IF sc_ship.a_ship(24)>1
*WAIT WINDOW "greater than 1" TIMEOUT 2
    
sc_ship.a_ship(63)="N"
ELSE
*WAIT WINDOW  "First package" TIMEOUT 1
SC_SHIP.A_CUS(24) = "Commercial"
ENDIF

ENDFUNC


*****************************************************************
FUNCTION AFTER_FSMOVE()
******************************************************************

**MAN.MSODBC=SFS THEN IT'S A SAP SHIPMENT THAT WAS FUTURE SHIPPED AND NEEDS TO BE UPDATED IN SAP
**FUTURE SHIPS MUST BE MOVED DURING OPEN DAY PROCEDURE.
 **CAN'T JUST RETRIEVE FROM FUTURE!!!!!!!

LNOLD = select()
WAIT WINDOW "Updating SAP with data, PLEASE WAIT " NOWAIT
*BROW &&IN SHIPPER
lcshipname = ALLTRIM(SHIPPER.manfil)
USE(lcShipName) IN 0 AGAIN SHARED ALIAS OUR_MSMAN

SELECT OUR_MSMAN

GO top
*BROW &&IN MAN 
SCAN WHILE .T.
 LOCATE FOR OUR_MSMAN.MSODBC="SFS" AND OUR_MSMAN.SHPSTS<>"V"
 
 **CHANGE OUR_MSMAN.MSODBC WITH "S" TO END THE PROCESS
 
 IF FOUND() 
* 	BROW &&FINDS IT
*!*	 SELE * FROM OUR_MSMAN WHERE atcc(lccurrinv, invnum) >0 and man.shpsts <> "V"  &&CAN'T USE CURSOR into cursor OUR_MSMAN
*!*		SELE OUR_MSMAN  &&NOT USING CURSOR, SHOULDN'T NEED TO SELECT AGAIN

*!*		GO TOP
*!*		SCAN

	IF LEN(ALLT(OUR_MSMAN.MSODBC)) > 0		&& CONTINUE WITH INTERFACE
	** UNIVERSAL DATA FOR BOTH INTERFACES
	lcO=""
	lcSycd=""
	lcId=""
	lcd1 = ""
	lcd2 = ""
	lcd3 = ""
	ldtotal = ""
	
		STORE allt(OUR_MSMAN.invnum) TO lcO	&& gcORDER
		store allt(GCSYCD) to lcsycd		&& system id 
		store allt(OUR_MSMAN.uniqueid) to lcID&& uniqueid
		store substr(dtoc(OUR_MSMAN.date),7,4) to lcd1
		store substr(dtoc(OUR_MSMAN.date),1,2) to lcd2
		store substr(dtoc(OUR_MSMAN.date),4,2) to lcd3
		ldtotal = lcd1+lcd2+lcd3
		lcd = val(ldTotal)

	store SUBSTR(OUR_MSMAN.time,1,2) + SUBSTR(OUR_MSMAN.time,4,2) + padl(ALLT(STR(SEC(DATETIME()))),2,"0")  to lctime	&&shiptime

** processed before  this should only be for gcsycd 1, Trade 
		store "" to lcProc	&& default always for Trumble (gcsycd 2) must put logic for trade side as trade is swept several times per day
				
		if glProc=.t.  and gcSycd = "1" && if package was processed and voided in TRade, set flag to "V" so SAP knows it is an update.  
				store "V" to lcProc
		endif
		
		if glProc=.f. and gcSycd="1" &&new record, no void
			store " " to lcProc 
         endif
          lntotfrt = 0   
        STORE OUR_MSMAN.PAYMENT TO LNPAY
		STORE allt(OUR_MSMAN.packer) TO lcP			    && apss logon
		store allt(str(OUR_MSMAN.curbox)) to lcCB			&& current box TEXT
		store OUR_MSMAN.curbox to lnCB					&& current box NUMERIC
		STORE ALLT(OUR_MSMAN.shipper)+ALLT(OUR_MSMAN.upstyp) TO lcCarr	&& ship via code
		store allt(str(OUR_MSMAN.maxbox)) to lcMB			&& current box
		store allt(str(OUR_MSMAN.pieces)) to lcPc
		STORE allt(OUR_MSMAN.zip) TO lcZp					&& postalcode
		STORE OUR_MSMAN.wight TO lnW		&& WEIGHT
		store allt(OUR_MSMAN.zone) to lcZn
		store allt(OUR_MSMAN.picknum) to lcPk
		STORE allt(OUR_MSMAN.trknum) TO lcT				&& 

		STORE OUR_MSMAN.SHIPCHG + OUR_MSMAN.genchg1 TO lntotfrt
			
		STORE OUR_MSMAN.SHIPCHG + OUR_MSMAN.genchg1 TO lnbase
											
		store allt(OUR_MSMAN.ucc128) to lcUcc

		store OUR_MSMAN.codchg to lncodchg		&&& amount of COD on package
		store OUR_MSMAN.CODfee to lnCod
		store OUR_MSMAN.dvi to lndvc	&&dvfee
		store OUR_MSMAN.dv to lndvia	&& dv amount
		store OUR_MSMAN.dvpip to lnaia	&& ALT /Pip dv amount
		store OUR_MSMAN.dvipip to lnaic	&& ALT/PIP INS FEE
		STORE OUR_MSMAN.AODFEE TO lnaodc	&& aod charges
		store OUR_MSMAN.tagfee to lnctc	&& call tag fee
		store OUR_MSMAN.hand to lnhf	&& handling fee

		IF allt(OUR_MSMAN.codf) = "Y"
			store "COD TYPE" + allt(str(OUR_MSMAN.anynum)) to lx	&& cod type for charges
		ELSE
			store "COD TYPE" to lx	&& non cod type for charges
		ENDIF

		STORE "" TO lcTmp1
		STORE "" TO lcTmp2
		STORE "" TO lcTmp3
		STORE "" TO lcTmp4
		STORE "" TO lcTmp5
		STORE "" TO lcTmp6
		STORE "" TO lcTmp7
		STORE "" TO lcTmp8
		STORE "" TO lcTmp9
		STORE "" TO lcTmp10
		STORE "" TO LCTMP11
		STORE "" TO lcTotal

**?lnbase,

        iF LNPAY=1
		STORE "MNSEQ, MNMODE, MNPKG, MNLTL, MNWGT, MNZONE, MNCAR, MNTRCK, MNTXT1, " TO lcTmp2
		store "MNBASE, MNUCC, MNCOD, MNDVC, MNDVIA, MNAIA, MNAIC, MNAODC, MNCTC, " to lctmp3
		store "MNHFEE, MNNET, MNCODC, MNPOST) " to lctmp4

		STORE "VALUES (?lcd, ?lcTime, ?lcO, ?lcsycd, ?lcID, ?lcP, ?lcCB, ?lcCarr, ?lcMB, ?lcPc, " TO lcTmp5
		store "?lnW, ?lcZn, ?lcPk, ?lcT, ?lx, " to lctmp6
		STORE "?lnbase, ?lcUcc, ?lncod, ?lndvc, ?lndvia, ?lnaia, ?lnaic, ?lnaodc, ?lnctc, " to lctmp7
		store "?lnhf, ?lntotfrt, ?lncodchg, ?lcproc)" to lctmp8
        
         ELSE  &&NOT PREPAID-SEND ZERO'S
         
         STORE "MNSEQ, MNMODE, MNPKG, MNLTL, MNWGT, MNZONE, MNCAR, MNTRCK, MNTXT1, " TO lcTmp2
		store "MNBASE, MNUCC, MNCOD, MNDVC, MNDVIA, MNAIA, MNAIC, MNAODC, MNCTC, " to lctmp3
		store "MNHFEE, MNNET, MNCODC, MNPOST) " to lctmp4
		
	STORE "VALUES (?lcd, ?lcTime, ?lcO, ?lcsycd, ?lcID, ?lcP, ?lcCB, ?lcCarr, ?lcMB, ?lcPc, " TO lcTmp5
		store "?lnW, ?lcZn, ?lcPk, ?lcT, ?lx, " to lctmp6
		STORE "0.00, ?lcUcc, ?lncod, ?lndvc, ?lndvia, ?lnaia, ?lnaic, ?lnaodc, ?lnctc, " to lctmp7
		store "?lnhf, 0.00, ?lncodchg, ?lcproc)" to lctmp8	

    ENDIF  &&lnpay=1
** END OF UNIVERSAL DATA
&& SAP ORDER
			
		IF GCSYCD == "1"	&& TRADE &&change because of intermittent problem
		*	wait window "trade" timeout .3&&TEST COMMENT OUT PRODUCTION 
			STORE "INSERT INTO PRD.prd.ZTRDAPSS(MNDTE, MNTIME, MNORD, MNSYCD, MNUNID, MNSYS, " TO lcTmp1
		ENDIF	
	
        IF GCSYCD == "2"	&& TRUMBLE EXACTLY &&change because of intermittent problem
       *	 wait window "TRUMBLE" timeout .3 &&TEST COMMENT OUT PRODUCTION 
			STORE "INSERT INTO PRD.prd.ZAPSS(MNDTE, MNTIME, MNORD, MNSYCD, MNUNID, MNSYS, " TO lcTmp1
		ENDIF	

		STORE lcTmp1+lcTmp2+lcTmp3+lcTmp4+lcTmp5+lcTmp6+lcTmp7+lcTmp8 TO lcTotal
	
		gnsqlerror2 = 0
** INSERT SQL STATEMENT
		IF ALLT(OUR_MSMAN.MSODBC)= "SFS" AND GNSQLSAP >0
			gnsqlerror=SQLEXEC(gnsqlSAP,lcTotal)
		REPLACE OUR_MSMAN.MSODBC WITH "SDF" &&TO END THE PROCESS SAP DAILY FILE
		ENDIF
*** END OF INSERT STATEMENT
		IF gnSqlError>0
			wait window "Saving record data.."+ALLTRIM(gcOrder) TIMEOUT .3  nowait
            		ELSE
			MESSAGEBOX("Bad insert result code for 'header' record..."+str(gnSqlError))
		select(lnold)
		RETURN .F.
 		ENDIF
*
		IF GNSQLSAP < 1
			messagebox("SAP ODBC connection not available."+chr(13)+"You must EXIT SPSS and Re-enter to reconnect." +;
			chr(13)+"Any shipments processed WILL NOT transfer information to SAP.", 0,"     ODBC CONNECTION NOT AVAILABLE")
		ENDIF
ELSE
   
	* MESSAGEBOX("MANUAL SHIPMENT, DATA NOT RETURNED TO INTERFACE")
			ENDIF	&& interfaced shipment, OUR_MSMAN.MSODBC O FOR AS400, S FOR SAP
	
		ENDIF && future shipment
	
	SELECT OUR_MSMAN
ENDSCAN

	IF USED('OUR_MSMAN')
		USE IN OUR_MSMAN
	ENDIF
		
		
**************************************
FUNCTION FUNCTIONAPPEND()
**************************************
**COPY INTO CUS_MAIN
**UNCOMMENT THE MISC POINTS 
**tParm = "PR8"

**RUN AT END OF DAY STUFF

lnOld=SELECT()
*IF SUBSTR(S_SHIPPER, 1,3)="PR8"

PUBLIC gcSPSSDIR , gcDHLCAN ,gnSwqty, gnSwnum, gnSw_Num2,gnTotOrd


gnswqty = 0 
gnswnum = 0 
GNSW_NUM2 = 0
GNTOTORD = 0

**APPEND_OURSWOG WITH INFO IN PR8MAN BASED ON PKGID'S
gcSPSSDIR = SYS(5) + SYS(2003)
gcDHLCAN = gcSPSSDIR + '\arc_dhlcan'
gcDatadir = gcSPSSDIR + '\DATA'
***first, put swog info into manswog as swog can have more records than man
WAIT WINDOW "CREATING FILE FOR DHL CANADA...PLEASE WAIT" TIMEOUT 4
SET SAFETY OFF
**FIRST FIND DATE FOR THE SHIPPER
IF ! USED('shipper')	
	USE 'shipper' ALIAS shipper
	ENDIF
	SELECT shipper
	lcdate=DTOC(shipper.date)
	
*	WAIT WINDOW "DATE = "+ALLTRIM(LCDATE)
**CLEAR OUT OLD DATA IN MANSWOG
if ! used('manswog')
	use 'manswog' alias manswog EXCLUSIVE	
	ZAP

	else
	endif
	
	IF USED ('MANSWOG')
	USE IN MANSWOG
	ENDIF
	

SET CENTURY on

	IF !USED ('MANSWOG')
	USE 'MANSWOG' IN 1 ALIAS MANSWOG
		ENDIF
	SELE MANSWOG
append from swog FOR ALLTRIM(lcdate)= DTOC(DATE)&&AND SHPSTS<>"V"

***next update with man info 	
	if ! used ('pr8man')
	use 'pr8man' in 2 alias man
	endif
		
	IF !USED ('MANSWOG')
	USE 'MANSWOG' IN 1 ALIAS MANSWOG
		ENDIF

	sele MANSWOG	

	
	*** scan through manswog; use man.dbf as a lookup table based on pkgid; 
*		replace all fields in manswog  with the value found in man 
*
SELE manswog
GO TOP

SCAN while !eof()
	
	*** lookup 'pkgid' in man( bould_deptsave.dbf)
	*
	SELE man
	GO TOP
	GCPKG=""
	GCPKG = ALLTRIM(MAN.PKGID)
	GNORDVAL = 0
	GNSW_NUM2 =0
	GNMSWIGHT = 0
	GNSQQTY = 0
	*WAIT WINDOW "PKG ID =  "+ALLTRIM(GCPKG)
	LOCATE FOR ALLT(man.pkgid)= ALLTRIM(SUBSTR(manswog.pkgid,1,10))AND SHPSTS<>"V"
	*
		*** update fields
		*
		IF FOUND()

			REPLACE manswog.company WITH alltr(man.company) 
			replace manswog.Csnam2 with  alltr(man.csnam2)
	 		replace manswog.address_1 with alltr(man.address)
			replace manswog.csadd2 with alltr(man.csadd2)
			replace manswog.Csadd3 with alltr(man.csadd3)
			replace manswog.Cscity with alltr(man.cscity)
			replace manswog.CsStat with alltr(man.csstat)
			replace manswog.Zip with alltr(man.zip)
			replace manswog.Phone with alltr(man.phone)
			replace manswog.Picknum with alltr(man.picknum)
			replace manswog.Airtrack with alltr(man.airtrack)
			replace manswog.POnum with alltr(man.ponum)
			replace manswog.wight with man.wight
			replace manswog.shpsts with alltr(man.shpsts)
			replace manswog.date with man.date
			replace manswog.sessionid with alltr(man.sessionid)
			replace manswog.refnum with alltr(man.refnum)
			replace manswog.dept with alltr(man.dept)	
			replace manswog.acctnum with "70243"
			REPLACE MANSWOG.CNTRYCODE WITH "CA"
			REPLACE MANSWOG.CNTRY WITH "CANADA"
			REPLACE MANSWOG.ORIG_CNTRY WITH "US"
			REPLACE MANSWOG.DESCRIP WITH "GREETING CARDS &/OR GIFT ITEMS"
			REPLACE MANSWOG.WINV WITH "N"
			REPLACE MANSWOG.EMAIL WITH ""
			REPLACE MANSWOG.SW_ANYTXT2 WITH ""
			
				
			*REPLACE MANSWOG.TOTPRIC WITH GNSW_NUM2
	*	SELECT * FROM MANSWOG INTO CURSOR OURMAN WHERE ALLTRIM(PKGID)=ALLTRIM(PARENT) AND SHPSTS<>"V"
				
		ENDIF
	
       SELE manswog	&& 
		ENDSCAN
		GO top
		SCAN WHILE ! EOF()
		IF EMPTY(manswog.company)

		DELETE
			ENDIF
		
		
		ENDSCAN
						
			SELE MANSWOG		
		GO TOP
	SCAN WHILE ! EOF()	
		
			gnswqty=MANswog.sw_qty 
 			gnswnum=manswog.sw_anynum &&column 22
			*GNSW_NUM2 = GNSW_num2
			GNSW_NUM2=(gnswqty * gnswnum)  &&QTY X COST	
		replace manswog.sw_anynum2 with gnsw_num2&& ,9,2
	 Sele * From manSWOG Into Cursor our_SWOG Where Allt(pkgid) = Allt(parent)&& And SHPSTS <> "V"
				
        Calc Sum(our_SWOG.SW_ANYNUM2)To GNTOTORD
	REPLACE manSWOG.SW_QTY2 WITH GNTOTORD FOR ALLTRIM(PKGID) = ALLTRIM(parent)
		ENDSCAN
				
		IF USED ('OUR_SWOG')
		USE IN OUR_SWOG
	ENDIF

					
	 &&FOR tRUMBLE DHL GLOBAL MAIL 
*!*			CALCULATE SUM (MANSWOG.SW_ANYNUM2) TO GNSW_NUM2
*!*				CALCULATE SUM (MANSWOG.WIGHT) TO GNMSWIGHT
*!*				CALCULATE SUM (MANSWOG.SW_QTY) TO GNSWQTY 
*!*					MESSAGEBOX ("RECORD FOLLOWING DATA FOR REPORTS", 0 + 16 , "BOL DATA" )
*!*			
*!*			MESSAGEBOX ("TOTAL # OF PIECES = " +ALLTRIM(STR(GNSWQTY,7,2)), 0 + 16 , "BOL DATA" )
*!*			MESSAGEBOX ("TOTAL WEIGHT = " +ALLTRIM(STR(GNMSWIGHT,7,2)), 0 + 16 , "BOL DATA" )
*!*			MESSAGEBOX ("TOTAL VALUE OF GOODS SHIPPED = " +ALLTRIM(STR(GNSW_NUM2,7,2)), 0 + 16 , "BOL DATA" )

if used ('MANswog')
	use in MANswog
	Endif
	
	IF USED ('man')
	USE IN man
	ENDIF
	
  	
	if ! used('manswog')
		use 'manswog' EXCL
			endif
		sele manswog
		PACK
IF FILE('70243')
	DELETE FILE 70243*
	endif

	LCFILENAME= ""
	LCFILENAME= "70243" + "_"+ ALLTRIM(STR(YEAR(DATE())))+ALLTRIM(STR(MONTH(DATE())))+ALLTRIM(STR(DAY(DATE()))) +"_PARCEL.DAT"
	
	
COPY TO ALLTRIM(LCFILENAME) fields ACCTNUM, PKGID, company, address_1, csadd2, csadd3, cscity, CSSTAT, ZIP, PHONE,email, CNTRYCODE, CNTRY, ORIG_CNTRY, DESCRIP, SW_QTY2, WIGHT, WINV,SW_QTY, CHILD, SW_DESC, SW_ANYNUM, AIRTRACK, SW_ANYTXT2  delimited with ' ' with character | 
COPY FILE alltrim(lcfilename) TO &gcDataDir

IF USED('manswog')
	USE IN manswog
	ENDIF
	
	COPY FILE ALLTRIM(lcFileName) TO &gcDHLCAN
	
	IF !USED ('PR8MAN')
		USE PR8MAN IN 0 ALIAS 'MAN'
		ENDIF
IF ! USED('SHIPPER')	
	USE SHIPPER IN 1 ALIAS 'SHIPPER'
	ENDIF
	
		
	RETURN 

*ENDIF &&PR8
SELECT (lnOld)