*&---------------------------------------------------------------------*
*& Report ZIT_ASSET_UPLOAD
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT ZIT_ASSET_UPLOAD.

TABLES: ANLA, ANLZ, ANLU, LFA1, icon.
TYPE-POOLS truxs.

" ======================================================== "

TYPES: BEGIN OF t_tab,
        comp_code TYPE ANLA-BUKRS,
        asset_id TYPE ANLA-ANLN1,
        sub_num TYPE ANLA-ANLN2,
        internal_no TYPE ANLA-INVNR,
        plant TYPE ANLZ-WERKS,
        location TYPE ANLZ-STORT,
        room TYPE ANLZ-RAUMN,
        brand TYPE ANLH-ANLHTXT,
        manufacturer TYPE ANLA-HERST,
        serial_num TYPE ANLA-SERNR,
        video_card TYPE ANLU-ZZVCARD,
        os_class TYPE ANLU-ZZOSCLASS,
        processor TYPE ANLU-ZZPROCESR,
        ram TYPE ANLU-ZZRAM,
        storage TYPE ANLU-ZZSTRGE,
        ospk TYPE ANLU-ZZOSPK,
        dept_use TYPE ANLU-ZZDEPT,
        emp_id TYPE ANLU-ZZLIFNR,
        ip_add TYPE ANLU-ZZIPADD,
        anydesk TYPE ANLU-ZZANYDESK,
        remarks TYPE string,
       END OF t_tab.

" ======================================================== "

TYPES: BEGIN OF upload_format,
        comp_code TYPE string,
        asset_id TYPE string,
        sub_num TYPE string,
        internal_no TYPE string,
        plant TYPE string,
        location TYPE anlz-stort,
        room TYPE anlz-raumn,
        brand TYPE string,
        manufacturer TYPE string,
        serial_num TYPE string,
        video_card TYPE string,
        os_class TYPE string,
        processor TYPE string,
        ram TYPE string,
        storage TYPE string,
        ospk TYPE string,
        dept_use TYPE string,
        emp_id TYPE string,
        ip_add TYPE string,
        anydesk TYPE string,
        remarks TYPE string,
       END OF upload_format.


TYPES: BEGIN OF ty_logs,
         status_icon TYPE icon-id,
         mat         TYPE string,
         msg         TYPE string,
       END OF ty_logs.

" ======================================================== "

DATA: t_upload TYPE STANDARD TABLE OF t_tab,
      wa_upload TYPE t_tab,
      lt_format TYPE STANDARD TABLE OF upload_format,
      wa_format TYPE upload_format,
      lt_main_upload TYPE  STANDARD TABLE OF t_tab,
      wa_main_upload TYPe t_tab,
      it_type TYPE truxs_t_text_data,
      c_file TYPE string,
      lv_counter TYPE i,

      lv_file_path TYPE string,
      lv_full_path TYPE string,
      lt_file_table TYPE TABLE OF string,
      lv_rc TYPE i,
      lv_workdir    TYPE string,
      lv_username   TYPE string,
      lv_loc_data   TYPE char4,
      lv_bukrs      TYPE string,
      lv_textlines  TYPE string,
      lt_textlines TYPE TABLE OF tline,
      ls_line TYPE tline,
      lv_error_found TYPE abap_bool VALUE abap_false.

" ======================================================== "

DATA: gs_logs TYPE ty_logs,
      gt_logs TYPE TABLE OF ty_logs,
      gt_anla TYPE TABLE OF ANLA,
      gs_anla TYPE ANLA,
      p_success TYPE c LENGTH 1,

      lt_error_msgs TYPE TABLE OF string,
      lt_success_msgs TYPE TABLE OF string,
      lt_no_change_msgs TYPE TABLE OF string,
      lv_error_msg  TYPE string,
      lv_success_msg  TYPE string,
      lv_no_change_msg TYPE string,
      lv_id TYPE icon-id,
      lv_concats TYPE c LENGTH 70,
      ls_header TYPE THEAD,
      lv_type TYPE char4,
      lv_temp TYPE string,
      lv_is_numeric TYPE abap_bool,
      str_comp_code TYPE string,
      str_asset_id TYPE string,
      str_sub_num TYPE string.

" ======================================================== "

DATA:   l_bukrs         TYPE bapi1022_1-comp_code,
        l_anln1         TYPE bapi1022_1-assetmaino,
        l_anln2         TYPE bapi1022_1-assetsubno,

        ls_gen          TYPE bapi1022_feglg001, "ANLA
        ls_genx         TYPE bapi1022_feglg001x,

        ls_origin       TYPE bapi1022_feglg009,
        ls_originx      TYPE bapi1022_feglg009x,

        ls_time         TYPE bapi1022_feglg003, "ANLZ
        ls_timex        TYPE bapi1022_feglg003x,

        ls_bapiret      TYPE bapiret2,

        lt_extensionin  TYPE TABLE OF bapiparex,
        ls_extensionin  LIKE LINE OF lt_extensionin,

        lt_bapi_anlu    TYPE TABLE OF bapi_te_anlu,
        ls_bapi_anlu    TYPE bapi_te_anlu, "for extension
        ls_anlu         TYPE anlu,

        l_asset_created TYPE bapi1022_reference.

*        lt_assetmasterdata TYPE bapi1022_asmd,
*        lt_assetmasterdatapx TYPE bapi1022_asmd_x,
*        ls_assetmasterdata TYPE bapi1022_asmd,
*        ls_assetmasterdatapx TYPE bapi1022_asmd_x.

" ======================================================== "

SELECTION-SCREEN BEGIN OF BLOCK s_block1 WITH FRAME TITLE TEXT-001.

  PARAMETERS p_file TYPE rlgrap-filename.

SELECTION-SCREEN END OF BLOCK s_block1.

" ======================================================== "

SELECTION-SCREEN: PUSHBUTTON /1(20) btn_dl USER-COMMAND download_template.

INITIALIZATION.
  btn_dl = 'Download Template'.


AT SELECTION-SCREEN.
  IF sy-ucomm = 'DOWNLOAD_TEMPLATE'.
    PERFORM dl_template.
  ENDIF.

" ======================================================== "

c_file = p_file.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  CALL FUNCTION 'F4_FILENAME'
     IMPORTING
       file_name = p_file.

" ======================================================== "

START-OF-SELECTION.

    IF c_file IS INITIAL.
      MESSAGE 'Please upload a file before proceeding.' TYPE 'E'.
    ENDIF.

   CALL METHOD cl_gui_frontend_services=>gui_upload
    EXPORTING
      filename            = c_file
      filetype            = 'ASC'
      has_field_separator = '|'
    CHANGING
      data_tab            = t_upload[].
   IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
              WITH sy-msgv1 sy-msgv2.
      EXIT.
   ENDIF.

    lv_counter = 0.
    LOOP AT t_upload INTO wa_upload.
      lv_counter = lv_counter + 1.
      IF lv_counter = 1.
        CONTINUE.
      ENDIF.

      CLEAR wa_main_upload.
        wa_main_upload-comp_code = wa_upload-comp_code.
        wa_main_upload-asset_id = wa_upload-asset_id.
        wa_main_upload-sub_num = wa_upload-sub_num.
        wa_main_upload-internal_no = wa_upload-internal_no.
        wa_main_upload-plant = wa_upload-plant.
        wa_main_upload-location = wa_upload-location.
        wa_main_upload-room = wa_upload-room.
        wa_main_upload-brand = wa_upload-brand.
        wa_main_upload-manufacturer = wa_upload-manufacturer.
        wa_main_upload-serial_num = wa_upload-serial_num.
        wa_main_upload-video_card = wa_upload-video_card.
        wa_main_upload-os_class = wa_upload-os_class.
        wa_main_upload-processor = wa_upload-processor.
        wa_main_upload-ram = wa_upload-ram.
        wa_main_upload-storage = wa_upload-storage.
        wa_main_upload-ospk = wa_upload-ospk.
        wa_main_upload-dept_use = wa_upload-dept_use.
        wa_main_upload-emp_id = wa_upload-emp_id.
        wa_main_upload-ip_add = wa_upload-ip_add.
        wa_main_upload-anydesk = wa_upload-anydesk.
        wa_main_upload-remarks = wa_upload-remarks.
      APPEND wa_main_upload TO lt_main_upload.
    ENDLOOP.

    t_upload = lt_main_upload.
    PERFORM asset_change.
    CLEAR lt_main_upload.

END-OF-SELECTION.

" ======================================================== "

FORM asset_change.

  LOOP AT t_upload INTO wa_upload.

    CLEAR: ls_gen, ls_genx, ls_time, ls_timex, ls_origin, ls_originx,
           ls_bapi_anlu, lt_extensionin.

    DATA(asset_id) = wa_upload-asset_id.
      CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
        EXPORTING
          input  = asset_id
        IMPORTING
          output = asset_id.

    DATA(sub_num) = wa_upload-sub_num.
      CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
        EXPORTING
          input  = sub_num
        IMPORTING
          output = sub_num.

    DATA(lv_loc) = wa_upload-location.

    CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
      EXPORTING
        input  = lv_loc
      IMPORTING
        output = lv_loc_data.

    l_bukrs = wa_upload-comp_code.
    l_anln1 = asset_id.
    l_anln2 = sub_num.

    ls_gen-invent_no = wa_upload-internal_no.

    ls_time-plant = wa_upload-plant.
    ls_time-location = lv_loc_data.
    ls_time-room = wa_upload-room.

    ls_gen-main_descript = wa_upload-brand.
    ls_origin-manufacturer = wa_upload-manufacturer.
    ls_gen-serial_no = wa_upload-serial_num.

    ls_bapi_anlu-comp_code = l_bukrs.
    ls_bapi_anlu-assetmaino = asset_id.
    ls_bapi_anlu-assetsubno = sub_num.

    ls_genx-invent_no = 'X'.
    ls_timex-plant = 'X'.
    ls_timex-location = 'X'.
    ls_timex-room = 'X'.
    ls_genx-main_descript = 'X'.
    ls_originx-manufacturer = 'X'.
    ls_genx-serial_no = 'X'.

    " ======================================================== "

    IF wa_upload-sub_num IS INITIAL.
      wa_upload-sub_num = '0000'.
    ELSE.
      wa_upload-sub_num = '0000'.
    ENDIF.

    " ======================================================== "
    ls_bapi_anlu-zzname1 = ''.
    ls_bapi_anlu-zzlifnr = wa_upload-comp_code.

    IF wa_upload-emp_id IS NOT INITIAL.
      DATA(lv_emp_id) = wa_upload-emp_id.
      CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
        EXPORTING
          input  = lv_emp_id
        IMPORTING
          output = lv_emp_id.

*      SELECT SINGLE name1, name2 FROM lfa1                                           " commented on 06/18/2025
*        INTO @DATA(ls_lfa1_zz)
*        WHERE lifnr = @lv_emp_id.

      SELECT SINGLE businesspartnerfullname FROM a_businesspartner                    " added on 06/18/2025
        INTO @DATA(ls_lfa1_zz)
        WHERE businesspartner = @lv_emp_id.

      IF sy-subrc = 0.
*        ls_bapi_anlu-zzname1 = ls_lfa1_zz-name1 && ' ' && ls_lfa1_zz-name2.          " commented on 06/18/2025
        ls_bapi_anlu-zzname1 = ls_lfa1_zz.
        ls_bapi_anlu-zzlifnr = wa_upload-emp_id.
      ELSE.
        ls_bapi_anlu-zzname1 = space.
        ls_bapi_anlu-zzlifnr = space.
      ENDIF.
    ELSE.
      ls_bapi_anlu-zzname1 = space.
      ls_bapi_anlu-zzlifnr = space.
    ENDIF.

    " ======================================================== "

    IF wa_upload-video_card IS NOT INITIAL.
      ls_bapi_anlu-zzvcard = wa_upload-video_card.
    ELSE.
      ls_bapi_anlu-zzvcard = space.
    ENDIF.

    " ======================================================== "

    IF wa_upload-os_class IS NOT INITIAL.
      ls_bapi_anlu-zzosclass = wa_upload-os_class.
    ELSE.
      ls_bapi_anlu-zzosclass = space.
    ENDIF.

    " ======================================================== "

    IF wa_upload-processor IS NOT INITIAL.
      ls_bapi_anlu-zzprocesr = wa_upload-processor.
    ELSE.
      ls_bapi_anlu-zzprocesr = space.
    ENDIF.

    " ======================================================== "

    IF wa_upload-ram IS NOT INITIAL.
      ls_bapi_anlu-zzram = wa_upload-ram.
    ELSE.
      ls_bapi_anlu-zzram = space.
    ENDIF.

    " ======================================================== "

    IF wa_upload-storage IS NOT INITIAL.
      ls_bapi_anlu-zzstrge = wa_upload-storage.
    ELSE.
      ls_bapi_anlu-zzstrge = space.
    ENDIF.

    " ======================================================== "

    IF wa_upload-ospk IS NOT INITIAL.
      ls_bapi_anlu-zzospk = wa_upload-ospk.
    ELSE.
      ls_bapi_anlu-zzospk = space.
    ENDIF.

    " ======================================================== "

    IF wa_upload-dept_use IS NOT INITIAL.
      ls_bapi_anlu-zzdept = 'X'.
    ELSE.
      ls_bapi_anlu-zzdept = space.
    ENDIF.

    " ======================================================== "

    IF wa_upload-ip_add IS NOT INITIAL.
      ls_bapi_anlu-zzipadd = wa_upload-ip_add.
    ELSE.
      ls_bapi_anlu-zzipadd = space.
    ENDIF.

    " ======================================================== "

    IF wa_upload-anydesk IS NOT INITIAL.
      ls_bapi_anlu-zzanydesk = wa_upload-anydesk.
    ELSE.
      ls_bapi_anlu-zzanydesk = space.
    ENDIF.

    " ======================================================== "

    MOVE 'BAPI_TE_ANLU' TO ls_extensionin-structure.

    CALL METHOD cl_abap_container_utilities=>fill_container_c
      EXPORTING
        im_value               = ls_bapi_anlu
      IMPORTING
        ex_container           = ls_extensionin+30
      EXCEPTIONS
        illegal_parameter_type = 1
        OTHERS                 = 2.

    APPEND ls_extensionin TO lt_extensionin.

    " ======================================================== "

    CALL FUNCTION 'BAPI_FIXEDASSET_CHANGE'
      EXPORTING
        companycode        = l_bukrs
        asset              = l_anln1
        subnumber          = l_anln2
        generaldata        = ls_gen
        generaldatax       = ls_genx
        timedependentdata  = ls_time
        timedependentdatax = ls_timex
        origin             = ls_origin
        originx            = ls_originx
      IMPORTING
        return             = ls_bapiret
      TABLES
        extensionin        = lt_extensionin.

    " ======================================================== "

    IF ls_bapiret-type EQ 'E'.
      gs_logs-status_icon = icon_led_red.
      gs_logs-mat = wa_upload-comp_code && ' ' && wa_upload-asset_id.
      gs_logs-msg = ls_bapiret-message.
      APPEND gs_logs TO gt_logs.

      CALL FUNCTION 'BAPI_TRANSACTION_ROLLBACK'.

      str_comp_code = wa_upload-comp_code.
      str_asset_id = wa_upload-asset_id.

      IF NOT str_comp_code CO '0123456789'.
        lv_error_msg = 'Error in Company Code'.
        APPEND lv_error_msg TO lt_error_msgs.
        EXIT.
      ENDIF.

      IF NOT str_asset_id CO '0123456789'.
        lv_error_msg = 'Error in Asset Number'.
        APPEND lv_error_msg TO lt_error_msgs.
        EXIT.
      ELSE.
        CONCATENATE 'Error in Asset'
                    wa_upload-comp_code
                    wa_upload-asset_id
                    ':'
                    ls_bapiret-message
                    INTO lv_error_msg
                    SEPARATED BY space.
        APPEND lv_error_msg TO lt_error_msgs.
        EXIT.
      ENDIF.

    ELSE.
      gs_logs-status_icon = icon_led_green.
      gs_logs-mat = wa_upload-comp_code && ' ' && wa_upload-asset_id.
      gs_logs-msg = ls_bapiret-message.
      APPEND gs_logs TO gt_logs.

      CALL FUNCTION 'BAPI_TRANSACTION_COMMIT'
        EXPORTING
          wait = 'X'.

      IF wa_upload-remarks IS NOT INITIAL.
         CLEAR ls_line.
         ls_line-tdline = wa_upload-remarks.
         ls_line-tdformat = '*'.
         APPEND ls_line TO lt_textlines.

         lv_bukrs = |{ wa_upload-comp_code }000000|.
         lv_concats = |{ lv_bukrs }{ wa_upload-asset_id }{ wa_upload-sub_num }|.
       ELSE.
         CLEAR wa_upload-remarks.
       ENDIF.

       ls_header-tdid    = 'XLTX'.
       ls_header-tdobject = 'ANLA'.
       ls_header-tdname  = lv_concats.
       ls_header-tdspras = sy-langu.

       CALL FUNCTION 'SAVE_TEXT'
         EXPORTING
           header          = ls_header
           savemode_direct = 'X'
         TABLES
           lines           = lt_textlines
         EXCEPTIONS
           id              = 1
           language        = 2
           name            = 3
           object          = 4
           others          = 5.

      IF sy-subrc = 0.
        COMMIT WORK.
        WRITE: 'Text saved successfully'.
      ELSE.
        WRITE: 'Error saving text. SY-SUBRC:', sy-subrc.
      ENDIF.

      IF ls_bapiret-message EQ 'No changes made'.
        lv_no_change_msg = 'No Changes made'.
        APPEND lv_no_change_msg TO lt_no_change_msgs.
      ELSE.
        CONCATENATE 'Successfully updated Asset' wa_upload-comp_code wa_upload-asset_id INTO lv_success_msg SEPARATED BY space.
        APPEND lv_success_msg TO lt_success_msgs.
      ENDIF.

    ENDIF.

  ENDLOOP.

   " ======================================================== "

  IF lt_error_msgs IS NOT INITIAL.
     MESSAGE lv_error_msg TYPE 'I'.
   ELSEIF  lt_success_msgs IS NOT INITIAL.
     MESSAGE lv_success_msg TYPE 'I'.
   ELSE.
     MESSAGE 'No Changes made' TYPE 'I'.
   ENDIF.

ENDFORM.

" ======================================================== "

FORM dl_template.

  wa_format-comp_code = 'Company Code'.
  wa_format-asset_id = 'Asset I.D'.
  wa_format-sub_num = 'Asset Sub-Number'.
  wa_format-internal_no = 'Internal Control no.'.
  wa_format-plant = 'Plant'.
  wa_format-location = 'Location'.
  wa_format-room = 'Room'.
  wa_format-brand  = 'Brand'.
  wa_format-manufacturer = 'Manufacturer'.
  wa_format-serial_num = 'Serial Number'.
  wa_format-video_card = 'Video Card'.
  wa_format-os_class = 'OS Classification'.
  wa_format-processor = 'Processor'.
  wa_format-ram = 'RAM'.
  wa_format-storage = 'Storage'.
  wa_format-ospk = 'OS Product Key'.
  wa_format-dept_use = 'Department Use?'.
  wa_format-emp_id = 'Employee I.D'.
  wa_format-ip_add = 'I.P Address'.
  wa_format-anydesk = 'Anydesk Address'.
  wa_format-remarks = 'Remarks'.
  APPEND wa_format TO lt_format.

  " ======================================================== "

  CALL METHOD cl_gui_frontend_services=>get_sapgui_workdir
    CHANGING
      sapworkdir = lv_workdir.

  FIND REGEX '\\Users\\([^\\]+)\\' IN lv_workdir SUBMATCHES lv_username.

  CONCATENATE 'C:\Users\' lv_username '\Downloads\IT_ASSET_TEMPLATE.xls' INTO lv_file_path.

  " ======================================================== "

    CALL METHOD cl_gui_frontend_services=>file_save_dialog
      EXPORTING
        window_title            = 'Save File As'
        default_extension       = 'xls'
        default_file_name       = 'IT_ASSET_TEMPLATE.xls'
      CHANGING
        filename                = lv_file_path
        path                    = lv_full_path
        fullpath                = lv_full_path
        user_action             = lv_rc.

  " ======================================================== "

       IF lv_rc = cl_gui_frontend_services=>action_ok.
        CALL FUNCTION 'GUI_DOWNLOAD'
           EXPORTING
             FILENAME                = lv_file_path
             FILETYPE                = 'ASC'
             write_field_separator   = 'X'
           TABLES
             DATA_TAB                = lt_format
           EXCEPTIONS
             file_write_error        = 1
             no_batch                = 2
             gui_refuse_filetransfer = 3
             invalid_type            = 4
             no_authority            = 5
             unknown_error           = 6
             header_not_allowed      = 7
             separator_not_allowed   = 8
             filesize_not_allowed    = 9
             header_too_long         = 10
             dp_error_create         = 11
             dp_error_send           = 12
             dp_error_write          = 13
             unknown_dp_error        = 14
             access_denied           = 15
             dp_out_of_memory        = 16
             disk_full               = 17
             dp_timeout              = 18
             file_not_found          = 19
             dataprovider_exception  = 20
             control_flush_error     = 21
             OTHERS                  = 22.

   " ======================================================== "

        IF sy-subrc = 0.
          MESSAGE 'Template downloaded successfully.' TYPE 'I'.
          CLEAR t_upload.
        ELSE.
          MESSAGE 'Failed to download template.' TYPE 'E'.
        ENDIF.
      ENDIF.

ENDFORM.
