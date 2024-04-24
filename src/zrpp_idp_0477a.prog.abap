*&---------------------------------------------------------------------*
*& Report ZRPP_IDP_0477
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT zrpp_idp_0477a.

TABLES: sscrfields,caufv.
DATA: lt_file        TYPE filetable,
      ls_file        LIKE LINE OF lt_file,
      lv_rc          TYPE i,
      lv_user_action TYPE i,
      lv_file_filter TYPE string.


TYPES: BEGIN OF t_input,
         werks     LIKE caufv-werks,
         aufnr     LIKE caufv-aufnr,
         prod_vers LIKE mkal-verid,
         plauf     LIKE caufv-gstrp,
         aufld     LIKE caufv-gltrp,
       END OF t_input.
DATA: input TYPE STANDARD TABLE OF t_input WITH HEADER LINE.
DATA: BEGIN OF t_log OCCURS 0,
        aufnr        LIKE caufv-aufnr,
        flag(1)      TYPE c,
        message(100) TYPE c,
      END OF t_log.
DATA: wo_log LIKE STANDARD TABLE OF zidppp081 WITH HEADER LINE.
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE txt01.
  PARAMETERS: p_werks LIKE aufk-werks OBLIGATORY DEFAULT 'ABG2',
              p_file  LIKE rlgrap-filename DEFAULT 'D:\zdppp477.xlsx' OBLIGATORY.
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN FUNCTION KEY 1.



INITIALIZATION.
  txt01 = '选择画面'.
  sscrfields-functxt_01 = '下载批导模板'.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  PERFORM open_file.

AT SELECTION-SCREEN.
  PERFORM download_template.


START-OF-SELECTION.
  PERFORM check_auth.
*  PERFORM upload_file.
  PERFORM upload_file1.
  PERFORM check_data.
  PERFORM process_data.
  PERFORM write_log.

*&---------------------------------------------------------------------*
*& Form open_file
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM open_file .
  CONCATENATE 'Excel (*.xls;*.xlsx)|*.xls;*.xlsx'
                  '|'
                  'All Files (*.*)|*.*'
                  INTO lv_file_filter.
  CALL METHOD cl_gui_frontend_services=>file_open_dialog
    EXPORTING
*     window_title            =
*     default_extension       =
*     default_filename        =
      file_filter             = lv_file_filter
*     with_encoding           =
      initial_directory       = 'D:\'
*     multiselection          =
    CHANGING
      file_table              = lt_file
      rc                      = lv_rc
      user_action             = lv_user_action
*     file_encoding           =
    EXCEPTIONS
      file_open_dialog_failed = 1
      cntl_error              = 2
      error_no_gui            = 3
      not_supported_by_gui    = 4
      OTHERS                  = 5.
  IF sy-subrc <> 0.
    MESSAGE 'File Open failed' TYPE 'E' RAISING error.      " File Open failed
  ENDIF.

  IF lv_user_action EQ cl_gui_frontend_services=>action_cancel.

    RETURN.
  ENDIF.

  READ TABLE lt_file INTO ls_file INDEX 1.
  IF sy-subrc = 0.
    p_file  = ls_file-filename.
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form upload_file
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM upload_file .
  DATA : lv_filename      TYPE string,
         lt_records       TYPE solix_tab,
         lv_headerxstring TYPE xstring,
         lv_filelength    TYPE i.

  FIELD-SYMBOLS : <gt_data>       TYPE STANDARD TABLE .
  FIELD-SYMBOLS : <ls_data>  TYPE any,
                  <lv_field> TYPE any.
  lv_filename = p_file.
  CALL METHOD cl_gui_frontend_services=>gui_upload
    EXPORTING
      filename                = lv_filename
      filetype                = 'BIN'
*     has_field_separator     = 'X'
*     header_length           = 0
*     read_by_line            = 'X'
*     dat_mode                = SPACE
    IMPORTING
      filelength              = lv_filelength
      header                  = lv_headerxstring
    CHANGING
      data_tab                = lt_records
    EXCEPTIONS
      file_open_error         = 1
      file_read_error         = 2
      no_batch                = 3
      gui_refuse_filetransfer = 4
      invalid_type            = 5
      no_authority            = 6
      unknown_error           = 7
      bad_data_format         = 8
      header_not_allowed      = 9
      separator_not_allowed   = 10
      header_too_long         = 11
      unknown_dp_error        = 12
      access_denied           = 13
      dp_out_of_memory        = 14
      disk_full               = 15
      dp_timeout              = 16
      not_supported_by_gui    = 17
      error_no_gui            = 18
      OTHERS                  = 19.
  CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
    EXPORTING
      input_length = lv_filelength
    IMPORTING
      buffer       = lv_headerxstring
    TABLES
      binary_tab   = lt_records
    EXCEPTIONS
      failed       = 1
      OTHERS       = 2.
  IF sy-subrc <> 0.
    "Implement suitable error handling here
  ENDIF.
  DATA : lo_excel_ref TYPE REF TO cl_fdt_xl_spreadsheet .

  TRY .
      lo_excel_ref = NEW cl_fdt_xl_spreadsheet(
                              document_name = lv_filename
                              xdocument     = lv_headerxstring ) .
    CATCH cx_fdt_excel_core.
      "Implement suitable error handling here
  ENDTRY .
  lo_excel_ref->if_fdt_doc_spreadsheet~get_worksheet_names(
   IMPORTING
     worksheet_names = DATA(lt_worksheets) ).

  IF NOT lt_worksheets IS INITIAL.
    READ TABLE lt_worksheets INTO DATA(lv_woksheetname) INDEX 1.

    DATA(lo_data_ref) = lo_excel_ref->if_fdt_doc_spreadsheet~get_itab_from_worksheet(
                                             lv_woksheetname ).
    "now you have excel work sheet data in dyanmic internal table
    ASSIGN lo_data_ref->* TO <gt_data>.
  ENDIF.

  DATA : lv_numberofcolumns   TYPE i,
         lv_date_string       TYPE string,
         lv_target_date_field TYPE datum,
         lt_dataset           TYPE TABLE OF t_input,
         ls_dataset           TYPE t_input.
  lv_numberofcolumns = 5.
  LOOP AT <gt_data> ASSIGNING <ls_data> FROM 2.
    CLEAR ls_dataset.
    DO lv_numberofcolumns TIMES.
      ASSIGN COMPONENT sy-index OF STRUCTURE <ls_data> TO <lv_field> .
      IF sy-subrc = 0.
        CASE sy-index.
          WHEN 1.
            ls_dataset-werks = <lv_field>.
          WHEN 2.
            ls_dataset-aufnr = <lv_field>.
          WHEN 3.
            ls_dataset-prod_vers = <lv_field>.

          WHEN 4.
            lv_date_string = <lv_field>.
            IF lv_date_string NS '/' AND lv_date_string NS '-'.
              ls_dataset-plauf = <lv_field>.
            ELSE.
              PERFORM date_convert USING lv_date_string.
              ls_dataset-plauf = lv_date_string.
            ENDIF.
          WHEN 5.
            lv_date_string = <lv_field>.
            IF lv_date_string NS '/' AND lv_date_string NS '-'.
              ls_dataset-aufld = <lv_field>.
            ELSE.
              PERFORM date_convert USING lv_date_string.
              ls_dataset-aufld = lv_date_string.
            ENDIF.
        ENDCASE.
      ENDIF.
    ENDDO.
    APPEND ls_dataset TO lt_dataset.
  ENDLOOP.
  input[] = lt_dataset[].

  REFRESH lt_dataset.
  IF input[] IS INITIAL.
    MESSAGE i003(zmm001).
    LEAVE LIST-PROCESSING.
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form date_convert
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*&      --> LV_DATE_STRING
*&---------------------------------------------------------------------*
FORM date_convert CHANGING iv_date_string.
  DATA: lv_convert_date(10) TYPE c.

  lv_convert_date = iv_date_string .
  DATA(lv_cnt) = strlen( lv_convert_date ).
  IF lv_cnt = 8.
    lv_convert_date = lv_convert_date+0(5) && '0' && lv_convert_date+5(2) && '0' && lv_convert_date+7(1).
  ENDIF.
  IF lv_cnt = 9.
    IF lv_convert_date+6(1) = '/' OR lv_convert_date+6(1) = '-'.
      lv_convert_date = lv_convert_date+0(5) && '0' && lv_convert_date+5(4).

    ELSEIF lv_convert_date+7(1) = '/' OR lv_convert_date+7(1) = '-'.
      lv_convert_date = lv_convert_date+0(8) && '0' && lv_convert_date+8(1).
    ENDIF.
  ENDIF.
  "date format YYYY/MM/DD
  FIND REGEX '^\d{4}[/|-]\d{1,2}[/|-]\d{1,2}$' IN lv_convert_date.
  IF sy-subrc = 0.
    CALL FUNCTION '/SAPDMC/LSM_DATE_CONVERT'
      EXPORTING
        date_in             = lv_convert_date
        date_format_in      = 'DYMD'
        to_output_format    = ''
        to_internal_format  = 'X'
      IMPORTING
        date_out            = lv_convert_date
      EXCEPTIONS
        illegal_date        = 1
        illegal_date_format = 2
        no_user_date_format = 3
        OTHERS              = 4.
  ELSE.

    " date format DD/MM/YYYY
    FIND REGEX '^\d{1,2}[/|-]\d{1,2}[/|-]\d{4}$' IN lv_convert_date.
    IF sy-subrc = 0.
      CALL FUNCTION '/SAPDMC/LSM_DATE_CONVERT'
        EXPORTING
          date_in             = lv_convert_date
          date_format_in      = 'DDMY'
          to_output_format    = ''
          to_internal_format  = 'X'
        IMPORTING
          date_out            = lv_convert_date
        EXCEPTIONS
          illegal_date        = 1
          illegal_date_format = 2
          no_user_date_format = 3
          OTHERS              = 4.
    ENDIF.

  ENDIF.

  IF sy-subrc = 0.
    iv_date_string = lv_convert_date .
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form download_template
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM download_template .

  DATA: lv_file_name         TYPE string,
        lv_path              TYPE string,
        lv_fullpath          TYPE string,
        lv_file_filter       TYPE string,
        lv_user_action       TYPE i,
        lv_default_file_name TYPE string.

  DATA: lv_file TYPE rlgrap-filename.
  DATA: lv_wwwdatatab TYPE wwwdatatab VALUE 'ZRPP_IDP_0477'.

  CONCATENATE 'Excel (*.xls;*.xlsx)|*.xls;*.xlsx'
                '|'
                'All Files (*.*)|*.*'
                INTO lv_file_filter.
  lv_default_file_name = '批量重展工单主数据的模板.xlsx'.
  CASE sscrfields-ucomm.
    WHEN 'FC01'.
      CALL METHOD cl_gui_frontend_services=>file_save_dialog
        EXPORTING
          window_title      = '批量重展工单主数据的模板'
*         default_extension =
          default_file_name = lv_default_file_name
*         with_encoding     =
          file_filter       = lv_file_filter
*         initial_directory =
*         prompt_on_overwrite  = 'X'
        CHANGING
          filename          = lv_file_name
          path              = lv_path
          fullpath          = lv_fullpath
          user_action       = lv_user_action
*         file_encoding     =
*    EXCEPTIONS
*         cntl_error        = 1
*         error_no_gui      = 2
*         not_supported_by_gui = 3
*         others            = 4
        .
      IF sy-subrc <> 0.
*   Implement suitable error handling here
      ENDIF.
      IF lv_user_action = cl_gui_frontend_services=>action_cancel.

      ELSE.
        lv_file = lv_fullpath.
        SELECT SINGLE * INTO CORRESPONDING FIELDS OF lv_wwwdatatab FROM wwwdata WHERE objid = 'ZRPP_IDP_0477' .

        "下载模板到指定路径
        CALL FUNCTION 'DOWNLOAD_WEB_OBJECT'
          EXPORTING
            key         = lv_wwwdatatab
            destination = lv_file.
*           IMPORTING
*           RC          =
*           CHANGING
*           TEMP        =
      ENDIF.

    WHEN OTHERS.
  ENDCASE.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form check_auth
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM check_auth .
  AUTHORITY-CHECK OBJECT 'C_AFKO_AWA'
          ID 'ACTVT' FIELD '02'
          ID 'WERKS' FIELD p_werks.
  IF sy-subrc NE 0.
    MESSAGE i000(pp) WITH '无工厂生产订单修改权限'.
    STOP.
  ENDIF.
  SELECT SINGLE @abap_true INTO @DATA(flag) FROM zidpcontrolvalue
     WHERE werks = @p_werks AND ctype = 'NPI_PLANT' AND indicator1 = 'ACT'.
  IF sy-subrc NE 0.
    MESSAGE i000(pp) WITH '非NPI工单,请再CO02执行'.
    STOP.
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form check_data
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM check_data .
  LOOP AT input ASSIGNING FIELD-SYMBOL(<fs_input>).
    IF <fs_input>-werks NE p_werks.
      MESSAGE i000(zpp) WITH <fs_input>-werks ' 工厂不一致,请确认!'.
      STOP.
    ENDIF.
    CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
      EXPORTING
        input  = <fs_input>-aufnr
      IMPORTING
        output = <fs_input>-aufnr.
    SELECT SINGLE substring( auart , 2 , 1 ) FROM caufv WHERE aufnr = @<fs_input>-aufnr INTO @DATA(l_type) .
    IF sy-subrc NE 0.
      MESSAGE i000(zpp) WITH <fs_input>-aufnr ' 工单号码不存在,请确认!'.
      STOP.
    ELSE.
      IF l_type = 'W' OR l_type = 'N'.
      ELSE.
        MESSAGE i000(zpp) WITH <fs_input>-aufnr ' 重工工单和拆解工单无法重展BOM,请确认!'.
        STOP.
      ENDIF.
    ENDIF.
    SELECT SINGLE werks FROM caufv WHERE aufnr = @<fs_input>-aufnr INTO @DATA(l_werks).
    IF l_werks NE p_werks.
      MESSAGE i000(zpp) WITH <fs_input>-aufnr ' 工单工厂与界面工厂不一致,请确认!'.
      STOP.
    ENDIF.
    SELECT SINGLE @abap_true INTO @DATA(exist) FROM resb
      WHERE aufnr = @<fs_input>-aufnr AND enmng > 0.
    IF sy-subrc EQ 0.
      MESSAGE i000(zpp) WITH <fs_input>-aufnr ' 工单存在已发料的料号无法重读,请确认!'.
      STOP.
    ENDIF.
  ENDLOOP.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form process_data
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM process_data .
  DATA: iporg    TYPE ni_nodeaddr,
        host(20) TYPE c.

  DATA: itab  LIKE bapi_pp_order_change,
        itabx LIKE bapi_pp_order_changex.
  DATA: errmsg LIKE bapiret2.
  LOOP AT input ASSIGNING FIELD-SYMBOL(<fs_input1>).
    itab-explosion_date = sy-datum.
    itab-explode_new = 'X'.
    itabx-explosion_date = 'X'.
    CALL FUNCTION 'BAPI_PRODORD_CHANGE'
      EXPORTING
        number     = <fs_input1>-aufnr
        orderdata  = itab
        orderdatax = itabx
      IMPORTING
        return     = errmsg
*       ORDER_TYPE =
*       ORDER_STATUS           =
*       MASTER_DATA_READ       =
*   TABLES
*       FSH_BUNDLES            =
      .
    t_log-aufnr = <fs_input1>-aufnr.
    IF errmsg-type = 'E' OR errmsg-type = 'A'.
      t_log-flag = 'E'.
      t_log-message = '失败!' && errmsg-message.
    ELSE.
      t_log-flag = 'S'.
      t_log-message = '成功!' && errmsg-message.

      wo_log-aufnr = <fs_input1>-aufnr.
      wo_log-filed = 'Read PP Master'.
      wo_log-aenam = sy-uname.
      wo_log-laeda = sy-datum.
      wo_log-times = sy-uzeit.
      wo_log-tcode = sy-tcode.
      wo_log-chnid = 'U'.
**  Get user IP,hostname
      CALL FUNCTION 'TH_USER_INFO'    " Get user IP,hostname
        IMPORTING
          addrstr  = iporg
          terminal = host
        EXCEPTIONS
          OTHERS   = 1.
      wo_log-hostip = iporg.
      wo_log-host = host.
      APPEND wo_log.

      INSERT zidppp081 FROM TABLE wo_log.
      COMMIT WORK AND WAIT.
      REFRESH wo_log.
    ENDIF.
    APPEND t_log.
  ENDLOOP.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form write_log
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM write_log .
  LOOP AT t_log WHERE flag = 'E'.
    WRITE: / t_log-aufnr , t_log-message COLOR 6.
  ENDLOOP.
  LOOP AT t_log WHERE flag = 'S'.
    WRITE: / t_log-aufnr , t_log-message COLOR 5.
  ENDLOOP.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form upload_file1
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM upload_file1 .
  DATA intern TYPE STANDARD TABLE OF zalsmex_tabline WITH HEADER LINE.
  FIELD-SYMBOLS:<fs>.
  CALL FUNCTION 'ZALSM_EXCEL_TO_INTERNAL_TABLE'
    EXPORTING
      filename                = p_file
      i_begin_col             = 1
      i_begin_row             = 2
      i_end_col               = 255
      i_end_row               = 65336
    TABLES
      intern                  = intern
    EXCEPTIONS
      inconsistent_parameters = 1
      upload_ole              = 2
      OTHERS                  = 3.

  IF sy-subrc <> 0.
    MESSAGE i368(00) WITH '上传失败,请调整模板'.
    STOP.
  ENDIF.
  IF intern[] IS INITIAL.
    MESSAGE i208(00) WITH 'No Data Upload'.
    STOP.
  ENDIF.
  LOOP AT intern.

    ASSIGN COMPONENT intern-col OF STRUCTURE input TO <fs>.
    MOVE intern-value TO <fs>.

    AT END OF zrow.
      APPEND input.CLEAR input.
    ENDAT.
  ENDLOOP.
ENDFORM.
