*___Class Definition_________________________________________________*

*___Class Main ______________________________________________________*
CLASS lcl_main DEFINITION .
  PUBLIC SECTION .
    TYPES :
        tt_first_conf_compliance TYPE STANDARD TABLE OF zaatp_s_first_conf_compliance WITH DEFAULT KEY.

    DATA :
      mt_data        TYPE tt_first_conf_compliance,
      mt_mappingdata TYPE tt_first_conf_compliance.

    DATA :
      gt_collect00 TYPE tt_first_conf_compliance, "genel
      gt_collect01 TYPE tt_first_conf_compliance, "kvgr2
      gt_collect02 TYPE tt_first_conf_compliance, "vbeln
      gt_collect03 TYPE tt_first_conf_compliance.."posnr

    DATA :
      mt_fname TYPE zcl_aatp_rpr_util=>tt_fname.

    DATA :
      mv_last_bopdate TYPE datum.

    DATA :
       mo_mdl TYPE REF TO zcl_aatp_main.

    METHODS :
      constructor,
      run,
      filter_data,
      calc_subtotals,
      export_list
        EXCEPTIONS
          error.

  PRIVATE SECTION.
    CONSTANTS : mc_red(4)    VALUE '@0A@',
                mc_yellow(4) VALUE '@09@',
                mc_green(4)  VALUE '@08@'.

    DATA : mv_impok.

    METHODS :
      _modify_data,
      _mapping_purchase,
      _mapping_sales,
      _get_line_color
        IMPORTING
          is_data         TYPE zaatp_s_first_conf_compliance
        RETURNING
          VALUE(rv_color) LIKE zcl_aatp_rpr_util=>c_color-red,
      _get_filename
        CHANGING
          cv_fullpath TYPE string
          cv_result   TYPE i,
      _download_xlsx
        IMPORTING
          iv_fullpath TYPE string
          iv_xlsx     TYPE xstring
        CHANGING
          cv_result   TYPE i,
      _prepxlsx
        RETURNING
          VALUE(rv_xlsx) TYPE xstring
        EXCEPTIONS
          error,
      _clear_lsdata
        CHANGING
          cs_data TYPE zaatp_s_first_conf_compliance,
      _delete_other_uuid_from_atplog
        IMPORTING
          is_data TYPE zaatp_s_first_conf_compliance.
ENDCLASS.

*__ Class Event _____________________________________________________*
CLASS lcl_event_handler DEFINITION .
  PUBLIC SECTION .
    DATA : mv_event(3).
    "SEL : selection screen
    "HDR : Header

    METHODS:
      constructor
        IMPORTING
          iv_event TYPE char03,

      handle_print_top_of_list FOR EVENT print_top_of_list OF cl_gui_alv_grid ,

      handle_user_command FOR EVENT user_command OF cl_gui_alv_grid
        IMPORTING e_ucomm,

      handle_toolbar FOR EVENT toolbar OF cl_gui_alv_grid
        IMPORTING e_object e_interactive,
      after_refresh FOR EVENT after_refresh OF cl_gui_alv_grid
        IMPORTING sender .

    METHODS :
      export_excel
        RETURNING
          VALUE(rv_xlsx) TYPE xstring.
ENDCLASS.


*__ Class View ______________________________________________________*
CLASS lcl_report_view DEFINITION .
  PUBLIC SECTION .
    CONSTANTS: c_str_h TYPE dd02l-tabname VALUE 'ZAATP_S_FIRST_CONF_COMPLIANCE'.

    CLASS-DATA : mo_view TYPE REF TO lcl_report_view.

    CLASS-METHODS :
      get_instance
        RETURNING VALUE(ro_view) TYPE REF TO lcl_report_view.

    " ALV DATA
    DATA: gt_fcat              TYPE lvc_t_fcat,
          gr_grid              TYPE REF TO cl_gui_alv_grid,
          gr_cont              TYPE REF TO cl_gui_custom_container,
          gs_toolbar_excluding TYPE ui_functions,
          gt_sort              TYPE lvc_t_sort.

    METHODS:
      display_data,
      refresh_alv
        IMPORTING
          ir_grid TYPE REF TO cl_gui_alv_grid,
      build_fcat IMPORTING i_str  TYPE dd02l-tabname
                 CHANGING  t_fcat TYPE lvc_t_fcat,
      change_subtotals.

  PRIVATE SECTION.

    METHODS:
      display_alv,
      exclude_functions,
      _set_sort.

ENDCLASS.

*___Class Implementation_____________________________________________*

*___Class Main ______________________________________________________*
CLASS lcl_main IMPLEMENTATION.
  METHOD constructor.
    "set colored columns
    mt_fname = VALUE #( ( fname = 'FIIL_SEVK_ADT' )
                        ( fname = 'FIIL_SEVK_TRH' )
                        ( fname = 'FIIL_SEVK_TTR' ) ).
  ENDMETHOD.

  METHOD run.

    IF p_kvgr IS NOT INITIAL.
      s_kvgr2[] = VALUE #( BASE s_kvgr2[] ( sign = 'I'
                                            option = 'NE'
                                            low = '' ) ).
    ENDIF.

    s_bsart[] = VALUE #( BASE s_bsart[] ( sign = 'I' option = 'EQ' low = 'ZS10' )
                                        ( sign = 'I' option = 'EQ' low = 'ZS12' )
                                        ( sign = 'I' option = 'EQ' low = 'ZSN1' )
                                        ( sign = 'I' option = 'EQ' low = 'ZSN3' ) ).


    mo_mdl = NEW zcl_aatp_main(
      ir_vbeln   = s_vbeln[]
      ir_posnr   = s_posnr[]
      ir_kunnr   = s_kunnr[]
      ir_matnr   = s_matnr[]
      ir_vkorg   = s_vkorg[]
      ir_vtweg   = s_vtweg[]
      ir_werks   = s_werks[]
      ir_auart   = s_auart[]
*      ir_sladate = s_sla[]
      ir_prodh   = s_prodh[]
      ir_matkl   = s_matkl[]
      ir_kvgr2   = s_kvgr2[]
      ir_mtart   = s_mtart[]
*      ir_lifsk   = s_lifsk[]
      ir_ebeln   = s_ebeln[]
      ir_ebelp   = s_ebelp[]
      ir_bsart   = s_bsart[]
      ir_facdate = s_facdt[]
*      ir_kerdat  = s_erdat[]
      ir_wadat   = s_wadat[]
    ).

    IF s_erdat-low IS INITIAL.
      s_erdat-low = '19000101'.
    ENDIF.
    IF s_erdat-high IS INITIAL.
      s_erdat-high = '99991231'.
    ENDIF.

    mv_last_bopdate = zcl_aatp_main=>get_last_bop_date( iv_datum = s_erdat-low ).

    IF mv_last_bopdate IS INITIAL.
      mo_mdl->set_kerdat( ir_erdat = VALUE #( ( sign = 'I'
                                                option = 'BT'
                                                low = s_erdat-low
                                                high = s_erdat-high ) ) ).
      mo_mdl->clear_atplog_key( ).
      IF p_purc EQ 'X'.
        mo_mdl->get_purchasing_data( ).
      ENDIF.
      IF p_sales EQ 'X'.
        mo_mdl->get_sales_data( ).
      ENDIF.
      mo_mdl->get_aatp_zlog_data( ).

    ELSE.
      "find pre bop data
      """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
      mo_mdl->set_kerdat( ir_erdat = VALUE #( ( sign = 'I'
                                                option = 'BT'
                                                low = s_erdat-low
                                                high = mv_last_bopdate ) ) ).
      mo_mdl->clear_atplog_key( ).
      IF p_purc EQ 'X'.
        mo_mdl->get_purchasing_data( ).
      ENDIF.
      IF p_sales EQ 'X'.
        mo_mdl->get_sales_data( ).
      ENDIF.
      mo_mdl->get_aatp_zlog_data( ).

      " find after bop data
      """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
      mo_mdl->set_kerdat( ir_erdat = VALUE #( ( sign = 'I'
                                                option = 'BT'
                                                low = CONV datum( mv_last_bopdate + 1 )
                                                high = s_erdat-high ) ) ).
      mo_mdl->clear_atplog_key( ).
      IF p_purc EQ 'X'.
        mo_mdl->get_purchasing_data( ).
      ENDIF.
      IF p_sales EQ 'X'.
        mo_mdl->get_sales_data( ).
      ENDIF.
      mo_mdl->get_aatp_log_data( ).
    ENDIF.

    mo_mdl->get_delivery_data(
      EXPORTING
        iv_wbstk    = 'C'
        iv_vbtp     = 'J' ).
    mo_mdl->get_sales_data_for_material( ).

    IF p_purc EQ 'X'.
      me->_mapping_purchase( ).
    ENDIF.
    IF p_sales EQ 'X'.
      me->_mapping_sales( ).
    ENDIF.

    me->_modify_data( ).
    me->filter_data( ).
    me->calc_subtotals( ).

  ENDMETHOD.

  METHOD _mapping_purchase.
    LOOP AT mo_mdl->mt_purc ASSIGNING FIELD-SYMBOL(<lfs_purc>).
      APPEND INITIAL LINE TO mt_mappingdata ASSIGNING FIELD-SYMBOL(<lfs_data>).
      <lfs_data>-vbeln            = <lfs_purc>-ebeln.
      <lfs_data>-posnr            = <lfs_purc>-ebelp_c6.
      <lfs_data>-kunnr            = <lfs_purc>-lifnr.
      <lfs_data>-matnr            = <lfs_purc>-matnr.
      <lfs_data>-vkorg            = <lfs_purc>-vkorg.
      <lfs_data>-vtweg            = <lfs_purc>-vtweg.
      <lfs_data>-werks            = <lfs_purc>-reswk.
      <lfs_data>-bsart            = <lfs_purc>-bsart.
      <lfs_data>-zzfactoryreqdate = <lfs_purc>-zzfactoryreqdate.
      <lfs_data>-zzchange_data    = <lfs_purc>-zzchange_date.
      <lfs_data>-erdat            = <lfs_purc>-creationdate.
      <lfs_data>-erzet            = <lfs_purc>-creationtime.
      <lfs_data>-kvgr2            = <lfs_purc>-kvgr2.
      <lfs_data>-bezei            = <lfs_purc>-bezei.
      <lfs_data>-matkl            = <lfs_purc>-matkl.
      <lfs_data>-wgbez            = <lfs_purc>-wgbez.
      <lfs_data>-kwmeng           = <lfs_purc>-menge.
      <lfs_data>-netwr            = <lfs_purc>-netpr.
      <lfs_data>-prodh            = <lfs_purc>-prodh.
      <lfs_data>-prodh_txt        = <lfs_purc>-vtext.
    ENDLOOP.
  ENDMETHOD.

  METHOD _mapping_sales.
    mo_mdl->get_sales_log_data( iv_fname = 'LIFSK' ).

    LOOP AT mo_mdl->mt_sales ASSIGNING FIELD-SYMBOL(<lfs_sales>).
      DATA(ls_cdlog) = VALUE #( mo_mdl->mt_cdlog[ objectid = <lfs_sales>-vbeln ] OPTIONAL ).

      APPEND INITIAL LINE TO mt_mappingdata ASSIGNING FIELD-SYMBOL(<lfs_data>).
      <lfs_data>-vbeln            = <lfs_sales>-vbeln.
      <lfs_data>-posnr            = <lfs_sales>-posnr.
      <lfs_data>-kunnr            = <lfs_sales>-kunnr.
      <lfs_data>-matnr            = <lfs_sales>-matnr.
      <lfs_data>-vkorg            = <lfs_sales>-vkorg.
      <lfs_data>-vtweg            = <lfs_sales>-vtweg.
      <lfs_data>-werks            = <lfs_sales>-werks.
      <lfs_data>-auart            = <lfs_sales>-auart.
      <lfs_data>-zzfactoryreqdate = <lfs_sales>-zzfactoryreqdate.
      <lfs_data>-zzchange_data    = <lfs_sales>-zzchange_date.
      <lfs_data>-erdat            = <lfs_sales>-erdat.
      <lfs_data>-erzet            = <lfs_sales>-erzet.
      <lfs_data>-kvgr2            = <lfs_sales>-kvgr2.
      <lfs_data>-bezei            = <lfs_sales>-bezei.
      <lfs_data>-matkl            = <lfs_sales>-matkl.
      <lfs_data>-wgbez            = <lfs_sales>-wgbez.
      <lfs_data>-kwmeng           = <lfs_sales>-kwmeng.
      <lfs_data>-netwr            = <lfs_sales>-netwr.
      <lfs_data>-prodh            = <lfs_sales>-prodh.
      <lfs_data>-prodh_txt        = <lfs_sales>-vtext.
      <lfs_data>-lifsk            = <lfs_sales>-lifsk.
      <lfs_data>-lifsk_date       = ls_cdlog-udate.

      CLEAR : ls_cdlog.
    ENDLOOP.
  ENDMETHOD.

  METHOD _modify_data.
    DATA :
      lv_datum1 TYPE datum,
      lv_datum2 TYPE datum.

    DATA :
      ls_data LIKE LINE OF mt_data.

    LOOP AT me->mt_mappingdata ASSIGNING FIELD-SYMBOL(<lfs_mapping>).

      ls_data-vbeln            = <lfs_mapping>-vbeln.
      ls_data-posnr            = <lfs_mapping>-posnr.
      ls_data-kunnr            = <lfs_mapping>-kunnr.
      ls_data-matnr            = <lfs_mapping>-matnr.
      ls_data-vkorg            = <lfs_mapping>-vkorg.
      ls_data-vtweg            = <lfs_mapping>-vtweg.
      ls_data-werks            = <lfs_mapping>-werks.
      ls_data-bsart            = <lfs_mapping>-bsart.
      ls_data-auart            = <lfs_mapping>-auart.
      ls_data-zzfactoryreqdate = <lfs_mapping>-zzfactoryreqdate.
      ls_data-zzchange_data    = <lfs_mapping>-zzchange_data.
      ls_data-erdat            = <lfs_mapping>-erdat.
      ls_data-erzet            = <lfs_mapping>-erzet.
      ls_data-kvgr2            = <lfs_mapping>-kvgr2.
      ls_data-bezei            = <lfs_mapping>-bezei.
      ls_data-matkl            = <lfs_mapping>-matkl.
      ls_data-wgbez            = <lfs_mapping>-wgbez.
      ls_data-kwmeng           = <lfs_mapping>-kwmeng.
      ls_data-netwr            = <lfs_mapping>-netwr.
      ls_data-prodh            = <lfs_mapping>-prodh.
      ls_data-prodh_txt        = <lfs_mapping>-prodh_txt.
      ls_data-lifsk            = <lfs_mapping>-lifsk.
      ls_data-lifsk_date       = <lfs_mapping>-lifsk_date.

      "1. add deliveries
      """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
      READ TABLE mo_mdl->mt_delivery TRANSPORTING NO FIELDS
        WITH KEY vgbel = <lfs_mapping>-vbeln
                 vgpos = <lfs_mapping>-posnr BINARY SEARCH.
      IF sy-subrc EQ 0.
        LOOP AT mo_mdl->mt_delivery ASSIGNING FIELD-SYMBOL(<lfs_delivery>)
          FROM sy-tabix
          WHERE vgbel EQ <lfs_mapping>-vbeln
            AND vgpos EQ <lfs_mapping>-posnr.

          READ TABLE mo_mdl->mt_atplog TRANSPORTING NO FIELDS
            WITH KEY atp_relevant_document      = <lfs_mapping>-vbeln
                     atp_relevant_document_item = <lfs_mapping>-posnr BINARY SEARCH.
          IF sy-subrc EQ 0.
            "1.1. add confirmed deliveries
            """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
            LOOP AT mo_mdl->mt_atplog ASSIGNING FIELD-SYMBOL(<lfs_atplog>)
                FROM sy-tabix
                WHERE atp_relevant_document      EQ <lfs_mapping>-vbeln
                  AND atp_relevant_document_item EQ <lfs_mapping>-posnr
                  AND confd_qty_after_run_inbaseunit GT 0
                  AND  abop_run_start_date_time GT zcl_aatp_rpr_util=>catt_add_to_time(
                                                       i_bdate    = COND #( WHEN <lfs_mapping>-zzchange_data IS NOT INITIAL THEN <lfs_mapping>-zzchange_data
                                                                            ELSE <lfs_mapping>-erdat )
                                                       i_btime    = <lfs_mapping>-erzet
                                                       i_add_hour = CONV #( p_xvalu )
                                                    ).
              IF <lfs_delivery>-lfimg EQ 0.
                EXIT.
              ENDIF.
              DATA(lv_add_flag) = 'X'.
              APPEND INITIAL LINE TO mt_data ASSIGNING FIELD-SYMBOL(<lfs_data>).
              <lfs_data> = ls_data.
              me->_clear_lsdata(  CHANGING cs_data = ls_data ).

              <lfs_data>-aboprunuuid = <lfs_atplog>-abop_run_uuid.
              me->_delete_other_uuid_from_atplog( is_data = <lfs_data> ).

              <lfs_data>-atprelevantdocscheduleline = <lfs_atplog>-atp_relevant_doc_schedule_line.

              CONVERT TIME STAMP <lfs_atplog>-confirmed_issue_date_time
                TIME ZONE sy-zonlo
                INTO DATE <lfs_data>-mal_cikis_trh
                     TIME DATA(lv_time).

              <lfs_data>-mal_cikis_trh = <lfs_data>-mal_cikis_trh - 1.
              <lfs_data>-fiil_sevk_trh = <lfs_delivery>-wadat_ist.

              IF <lfs_delivery>-lfimg LT <lfs_atplog>-confd_qty_after_run_inbaseunit.
                <lfs_data>-fiil_sevk_adt              =  <lfs_delivery>-lfimg.
                <lfs_data>-confdqtyafterruninbaseunit = <lfs_delivery>-lfimg.
              ELSE.
                <lfs_data>-fiil_sevk_adt =  <lfs_atplog>-confd_qty_after_run_inbaseunit.
                <lfs_data>-confdqtyafterruninbaseunit = <lfs_data>-fiil_sevk_adt.
              ENDIF.

              <lfs_atplog>-confd_qty_after_run_inbaseunit = <lfs_atplog>-confd_qty_after_run_inbaseunit
                                                            - <lfs_data>-fiil_sevk_adt.
              <lfs_delivery>-lfimg = <lfs_delivery>-lfimg
                                     - <lfs_data>-fiil_sevk_adt.
              TRY.
                  <lfs_data>-fiil_sevk_ttr = ( <lfs_mapping>-netwr / <lfs_mapping>-kwmeng )
                                             * <lfs_data>-fiil_sevk_adt.
                  <lfs_data>-teyit_ttr = ( <lfs_mapping>-netwr / <lfs_mapping>-kwmeng )
                                          * <lfs_data>-confdqtyafterruninbaseunit.
                  <lfs_data>-oran = <lfs_data>-fiil_sevk_adt / <lfs_data>-confdqtyafterruninbaseunit.
                CATCH cx_sy_zerodivide.
              ENDTRY.


              DATA(lv_color) = _get_line_color( <lfs_data> ).

              <lfs_data>-oran = COND #( WHEN lv_color EQ zcl_aatp_rpr_util=>c_color-red
                                         THEN 0
                                         ELSE 1 ).

              <lfs_data>-color
                = zcl_aatp_rpr_util=>alv_modify_color( iv_color = lv_color
                                                       it_fname = mt_fname ).

              <lfs_data>-adetsel = <lfs_data>-oran * <lfs_data>-fiil_sevk_adt.
              <lfs_data>-tutarsal = <lfs_data>-oran * <lfs_data>-fiil_sevk_ttr.

              CLEAR : lv_color.
            ENDLOOP.
          ENDIF.
          "1.2. add unconfirmed deliveries
          """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
          IF <lfs_delivery>-lfimg NE 0.
            lv_add_flag = 'X'.
            APPEND INITIAL LINE TO mt_data ASSIGNING <lfs_data>.
            <lfs_data> = ls_data.
            me->_clear_lsdata(  CHANGING cs_data = ls_data ).

            <lfs_data>-fiil_sevk_adt = <lfs_delivery>-lfimg.
            <lfs_data>-fiil_sevk_trh = <lfs_delivery>-wadat_ist.

            TRY.
                <lfs_data>-fiil_sevk_ttr = ( <lfs_mapping>-netwr / <lfs_mapping>-kwmeng )
                                           * <lfs_data>-fiil_sevk_adt.
                <lfs_data>-teyit_ttr = ( <lfs_mapping>-netwr / <lfs_mapping>-kwmeng )
                                       * <lfs_data>-confdqtyafterruninbaseunit.
              CATCH cx_sy_zerodivide.
            ENDTRY.

            lv_color = _get_line_color( <lfs_data> ).
            <lfs_data>-color
              = zcl_aatp_rpr_util=>alv_modify_color( iv_color = lv_color
                                                     it_fname = mt_fname ).
            <lfs_data>-oran = COND #( WHEN lv_color EQ zcl_aatp_rpr_util=>c_color-red
                                       THEN 0
                                       ELSE 1 ).
            <lfs_data>-adetsel = <lfs_data>-oran * <lfs_data>-fiil_sevk_adt.
            <lfs_data>-tutarsal = <lfs_data>-oran * <lfs_data>-fiil_sevk_ttr.
          ENDIF.
        ENDLOOP.
      ENDIF.

      "2. add undelivered confirmations
      """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
      READ TABLE mo_mdl->mt_atplog TRANSPORTING NO FIELDS
          WITH KEY atp_relevant_document      = <lfs_mapping>-vbeln
                   atp_relevant_document_item = <lfs_mapping>-posnr BINARY SEARCH.
      IF sy-subrc EQ 0.
        LOOP AT mo_mdl->mt_atplog ASSIGNING <lfs_atplog>
              FROM sy-tabix
              WHERE atp_relevant_document      EQ <lfs_mapping>-vbeln
                AND atp_relevant_document_item EQ <lfs_mapping>-posnr
                AND confd_qty_after_run_inbaseunit GT 0
                AND  abop_run_start_date_time GT zcl_aatp_rpr_util=>catt_add_to_time(
                                                     i_bdate    = COND #( WHEN <lfs_mapping>-zzchange_data IS NOT INITIAL THEN <lfs_mapping>-zzchange_data
                                                                            ELSE <lfs_mapping>-erdat )
                                                     i_btime    = <lfs_mapping>-erzet
                                                     i_add_hour = CONV #( p_xvalu )
                                                  ).
          lv_add_flag = 'X'.
          APPEND INITIAL LINE TO mt_data ASSIGNING <lfs_data>.
          <lfs_data> = ls_data.
          me->_clear_lsdata( CHANGING cs_data = ls_data ).

          <lfs_data>-aboprunuuid = <lfs_atplog>-abop_run_uuid.
          me->_delete_other_uuid_from_atplog( is_data = <lfs_data> ).

          <lfs_data>-atprelevantdocscheduleline = <lfs_atplog>-atp_relevant_doc_schedule_line.

          CONVERT TIME STAMP <lfs_atplog>-confirmed_issue_date_time
            TIME ZONE sy-zonlo
            INTO DATE <lfs_data>-mal_cikis_trh
                 TIME lv_time.

          <lfs_data>-mal_cikis_trh = <lfs_data>-mal_cikis_trh - 1.

          <lfs_data>-confdqtyafterruninbaseunit = <lfs_atplog>-confd_qty_after_run_inbaseunit.


          TRY.
              <lfs_data>-teyit_ttr = ( <lfs_mapping>-netwr / <lfs_mapping>-kwmeng )
                                      * <lfs_data>-confdqtyafterruninbaseunit.
            CATCH cx_sy_zerodivide.
          ENDTRY.


          lv_color = _get_line_color( <lfs_data> ).

          <lfs_data>-oran = COND #( WHEN lv_color EQ zcl_aatp_rpr_util=>c_color-red
                                     THEN 0
                                     ELSE 1 ).

          <lfs_data>-color
            = zcl_aatp_rpr_util=>alv_modify_color( iv_color = lv_color
                                                   it_fname = mt_fname ).

          <lfs_data>-adetsel = <lfs_data>-oran * <lfs_data>-fiil_sevk_adt.
          <lfs_data>-tutarsal = <lfs_data>-oran * <lfs_data>-fiil_sevk_ttr.

          CLEAR : lv_color.

        ENDLOOP.
      ENDIF.

      "3. teyit veya teslimatı yoksa yine de ekle
      """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
      IF lv_add_flag EQ abap_false.
        APPEND ls_data TO me->mt_data.
      ENDIF.

      CLEAR : ls_data, lv_color, lv_add_flag.
    ENDLOOP.

  ENDMETHOD.

  METHOD _get_line_color.
    DATA lv_datum1 TYPE datum.
    DATA lv_datum2 TYPE datum.

    lv_datum1 = is_data-mal_cikis_trh + p_yvalu.
    lv_datum2 = is_data-zzfactoryreqdate + p_yvalu.

    IF is_data-fiil_sevk_trh IS INITIAL .
      rv_color  = zcl_aatp_rpr_util=>c_color-red.
    ELSE.
      IF is_data-fiil_sevk_trh GT lv_datum1.
        IF is_data-fiil_sevk_trh GT lv_datum2.
          rv_color = zcl_aatp_rpr_util=>c_color-red.
        ELSE.
          rv_color = zcl_aatp_rpr_util=>c_color-green.
        ENDIF.
      ELSE.
        rv_color = zcl_aatp_rpr_util=>c_color-green.
      ENDIF.
    ENDIF.

  ENDMETHOD.

  METHOD calc_subtotals.
    SORT mt_data BY kvgr2 vbeln posnr.

    APPEND INITIAL LINE TO gt_collect00 ASSIGNING FIELD-SYMBOL(<lfs_collect00>).

    LOOP AT mt_data INTO DATA(lg_kvgr2) GROUP BY lg_kvgr2-kvgr2.

      APPEND INITIAL LINE TO gt_collect01 ASSIGNING FIELD-SYMBOL(<lfs_collect01>).
      <lfs_collect01>-kvgr2 = lg_kvgr2-kvgr2.


      LOOP AT GROUP lg_kvgr2 INTO DATA(ls_vbeln) GROUP BY ls_vbeln-vbeln.

        APPEND INITIAL LINE TO gt_collect02 ASSIGNING FIELD-SYMBOL(<lfs_collect02>).
        <lfs_collect02>-kvgr2 = ls_vbeln-kvgr2.
        <lfs_collect02>-vbeln = ls_vbeln-vbeln.

        LOOP AT GROUP ls_vbeln INTO DATA(ls_posnr) GROUP BY ls_posnr-posnr.
          APPEND INITIAL LINE TO gt_collect03 ASSIGNING FIELD-SYMBOL(<lfs_collect03>).
          <lfs_collect03>-kvgr2 = ls_posnr-kvgr2.
          <lfs_collect03>-vbeln = ls_posnr-vbeln.
          <lfs_collect03>-posnr = ls_posnr-posnr.

          LOOP AT GROUP ls_posnr INTO DATA(ls_data).
            <lfs_collect03>-kwmeng = <lfs_collect03>-kwmeng + ls_data-kwmeng.
            <lfs_collect03>-netwr = <lfs_collect03>-netwr + ls_data-netwr.
            <lfs_collect03>-teyit_ttr = <lfs_collect03>-teyit_ttr + ls_data-teyit_ttr.
            <lfs_collect03>-fiil_sevk_adt = <lfs_collect03>-fiil_sevk_adt + ls_data-fiil_sevk_adt.
            <lfs_collect03>-fiil_sevk_ttr = <lfs_collect03>-fiil_sevk_ttr + ls_data-fiil_sevk_ttr.
            <lfs_collect03>-adetsel = <lfs_collect03>-adetsel + ls_data-adetsel.
            <lfs_collect03>-tutarsal = <lfs_collect03>-tutarsal + ls_data-tutarsal.
            <lfs_collect03>-confdqtyafterruninbaseunit = <lfs_collect03>-confdqtyafterruninbaseunit + ls_data-confdqtyafterruninbaseunit.
          ENDLOOP.

          <lfs_collect02>-kwmeng        = <lfs_collect02>-kwmeng        + <lfs_collect03>-kwmeng.
          <lfs_collect02>-netwr        = <lfs_collect02>-netwr        + <lfs_collect03>-netwr.
          <lfs_collect02>-teyit_ttr     = <lfs_collect02>-teyit_ttr     + <lfs_collect03>-teyit_ttr.
          <lfs_collect02>-fiil_sevk_adt = <lfs_collect02>-fiil_sevk_adt + <lfs_collect03>-fiil_sevk_adt.
          <lfs_collect02>-fiil_sevk_ttr = <lfs_collect02>-fiil_sevk_ttr + <lfs_collect03>-fiil_sevk_ttr.
          <lfs_collect02>-adetsel       = <lfs_collect02>-adetsel       + <lfs_collect03>-adetsel.
          <lfs_collect02>-tutarsal      = <lfs_collect02>-tutarsal      + <lfs_collect03>-tutarsal.
          <lfs_collect02>-confdqtyafterruninbaseunit = <lfs_collect02>-confdqtyafterruninbaseunit + <lfs_collect03>-confdqtyafterruninbaseunit.

          "ortalamalar
          IF <lfs_collect03>-kwmeng NE 0.
            <lfs_collect03>-adetsel = <lfs_collect03>-adetsel / <lfs_collect03>-kwmeng.
          ELSE.
            <lfs_collect03>-adetsel = 0.
          ENDIF.

          IF <lfs_collect03>-netwr NE 0.
            <lfs_collect03>-tutarsal = <lfs_collect03>-tutarsal / <lfs_collect03>-netwr.
          ELSE.
            <lfs_collect03>-tutarsal = 0.
          ENDIF.

        ENDLOOP.

        <lfs_collect01>-kwmeng        = <lfs_collect01>-kwmeng        + <lfs_collect02>-kwmeng.
        <lfs_collect01>-netwr = <lfs_collect01>-netwr + <lfs_collect02>-netwr.
        <lfs_collect01>-teyit_ttr     = <lfs_collect01>-teyit_ttr     + <lfs_collect02>-teyit_ttr.
        <lfs_collect01>-fiil_sevk_adt = <lfs_collect01>-fiil_sevk_adt + <lfs_collect02>-fiil_sevk_adt.
        <lfs_collect01>-fiil_sevk_ttr = <lfs_collect01>-fiil_sevk_ttr + <lfs_collect02>-fiil_sevk_ttr.
        <lfs_collect01>-adetsel       = <lfs_collect01>-adetsel       + <lfs_collect02>-adetsel.
        <lfs_collect01>-tutarsal       = <lfs_collect01>-tutarsal       + <lfs_collect02>-tutarsal.
        <lfs_collect01>-confdqtyafterruninbaseunit = <lfs_collect01>-confdqtyafterruninbaseunit + <lfs_collect02>-confdqtyafterruninbaseunit.

        "ortalamalar
        IF <lfs_collect02>-netwr  NE 0.
          <lfs_collect02>-tutarsal = <lfs_collect02>-tutarsal / <lfs_collect02>-netwr.
        ELSE.
          <lfs_collect02>-tutarsal = 0.
        ENDIF.

        IF <lfs_collect02>-kwmeng NE 0.
          <lfs_collect02>-adetsel = <lfs_collect02>-adetsel / <lfs_collect02>-kwmeng.
        ELSE.
          <lfs_collect02>-adetsel = 0.
        ENDIF.

      ENDLOOP.

      <lfs_collect00>-kwmeng        = <lfs_collect00>-kwmeng        + <lfs_collect01>-kwmeng.
      <lfs_collect00>-netwr = <lfs_collect00>-netwr + <lfs_collect01>-netwr.
      <lfs_collect00>-teyit_ttr     = <lfs_collect00>-teyit_ttr     + <lfs_collect01>-teyit_ttr.
      <lfs_collect00>-fiil_sevk_adt = <lfs_collect00>-fiil_sevk_adt + <lfs_collect01>-fiil_sevk_adt.
      <lfs_collect00>-fiil_sevk_ttr = <lfs_collect00>-fiil_sevk_ttr + <lfs_collect01>-fiil_sevk_ttr.
      <lfs_collect00>-adetsel       = <lfs_collect00>-adetsel       + <lfs_collect01>-adetsel.
      <lfs_collect00>-tutarsal       = <lfs_collect00>-tutarsal        + <lfs_collect01>-tutarsal  .
      <lfs_collect00>-confdqtyafterruninbaseunit = <lfs_collect00>-confdqtyafterruninbaseunit + <lfs_collect01>-confdqtyafterruninbaseunit.

      "ortalamalar
      IF <lfs_collect01>-netwr NE 0.
        <lfs_collect01>-tutarsal = <lfs_collect01>-tutarsal / <lfs_collect01>-netwr .
      ELSE.
        <lfs_collect01>-tutarsal = 0.
      ENDIF.

      IF <lfs_collect01>-kwmeng NE 0.
        <lfs_collect01>-adetsel = <lfs_collect01>-adetsel / <lfs_collect01>-kwmeng  .
      ELSE.
        <lfs_collect01>-adetsel = 0.
      ENDIF.

    ENDLOOP.

    "ortalamalar
    IF <lfs_collect00>-netwr NE 0.
      <lfs_collect00>-tutarsal = <lfs_collect00>-tutarsal / <lfs_collect00>-netwr .
    ELSE.
      <lfs_collect00>-tutarsal = 0.
    ENDIF.

    IF <lfs_collect00>-kwmeng NE 0.
      <lfs_collect00>-adetsel = <lfs_collect00>-adetsel / <lfs_collect00>-kwmeng  .
    ELSE.
      <lfs_collect00>-adetsel = 0.
    ENDIF.

  ENDMETHOD.

  METHOD filter_data.
    DELETE mt_data WHERE fiil_sevk_trh NOT IN s_wadat[].
  ENDMETHOD.

  METHOD export_list.
    DATA(lo_conv_excel) = NEW zcl_bc_alv_to_excel( ).

    lo_conv_excel->get_alv_data(
      EXPORTING
        ir_grid = gr_report_view->gr_grid
      IMPORTING
        et_data = DATA(lt_data)
        et_info = DATA(lt_info)
    ).

    CALL FUNCTION 'ZAATP_FM_ALV_TO_EXCEL'
        IN BACKGROUND TASK
        AS SEPARATE UNIT
        DESTINATION 'NONE'
      EXPORTING
        it_data = lt_data
        it_info = lt_info
        iv_option = zcl_bc_excel_outopt=>c_mail
        iv_title  = sy-title
      .
    COMMIT WORK.
  ENDMETHOD.

  METHOD _get_filename.
    DATA : lv_path     TYPE string,
           lv_filename TYPE string.

    cl_gui_frontend_services=>get_desktop_directory( CHANGING desktop_directory = lv_path ).
    cl_gui_cfw=>flush( ).

    CONCATENATE sy-title '_' sy-datum '.XLSX' INTO lv_filename.

    CALL METHOD cl_gui_frontend_services=>file_save_dialog
      EXPORTING
        window_title      = 'Export Excel'
        default_extension = 'XLSX'
        file_filter       = 'Excel dosyası (*.XLSX)'
        default_file_name = lv_filename
        initial_directory = lv_path
      CHANGING
        filename          = lv_filename
        path              = lv_path
        fullpath          = cv_fullpath
        user_action       = cv_result.
  ENDMETHOD.

  METHOD _download_xlsx.
    DATA: lv_bytecount TYPE i,
          lt_rawdata   TYPE solix_tab.
    lt_rawdata = cl_bcs_convert=>xstring_to_solix( iv_xstring  = iv_xlsx ).
    lv_bytecount = xstrlen( iv_xlsx ).
    cl_gui_frontend_services=>gui_download(
      EXPORTING
        bin_filesize              = lv_bytecount
        filename                  = iv_fullpath
        filetype                  = 'BIN'
        no_auth_check             = abap_true
      CHANGING
        data_tab                  = lt_rawdata
      EXCEPTIONS
        file_write_error          = 1
        no_batch                  = 2
        gui_refuse_filetransfer   = 3
        invalid_type              = 4
        no_authority              = 5
        unknown_error             = 6
        header_not_allowed        = 7
        separator_not_allowed     = 8
        filesize_not_allowed      = 9
        header_too_long           = 10
        dp_error_create           = 11
        dp_error_send             = 12
        dp_error_write            = 13
        unknown_dp_error          = 14
        access_denied             = 15
        dp_out_of_memory          = 16
        disk_full                 = 17
        dp_timeout                = 18
        file_not_found            = 19
        dataprovider_exception    = 20
        control_flush_error       = 21
        not_supported_by_gui      = 22
        error_no_gui              = 23
        OTHERS                    = 24
    ).
    cv_result = sy-subrc.
    IF cv_result <> 0.
      MESSAGE 'File couldnt be saved!' TYPE 'I' DISPLAY LIKE 'E'.
    ENDIF.
  ENDMETHOD.

  METHOD _prepxlsx.
  ENDMETHOD.



  METHOD _clear_lsdata.
    CLEAR :
        cs_data-confdqtyafterruninbaseunit,
        cs_data-kwmeng,
        cs_data-mal_cikis_trh,
        cs_data-netwr.
  ENDMETHOD.


  METHOD _delete_other_uuid_from_atplog.
    "ilk bulduğumuz bop ile çalışılacak.
    DELETE mo_mdl->mt_atplog WHERE abop_run_uuid NE  is_data-aboprunuuid
                               AND atp_relevant_document EQ is_data-vbeln
                               AND atp_relevant_document_item EQ is_data-posnr.

  ENDMETHOD.

ENDCLASS.


*__ Class Event _____________________________________________________*
CLASS lcl_event_handler IMPLEMENTATION.

  METHOD : constructor.
    mv_event = iv_event.
  ENDMETHOD.

  METHOD handle_print_top_of_list .
  ENDMETHOD .

  METHOD handle_user_command.
    CASE e_ucomm.
      WHEN '&RNT'.
      WHEN '&F03' OR '&F12' OR '&F15' .
        LEAVE TO SCREEN 0.
      WHEN 'EXPO'.
        gr_main->export_list( ).
        gr_report_view->refresh_alv( ir_grid = gr_report_view->gr_grid ).
    ENDCASE.
  ENDMETHOD.

  METHOD handle_toolbar.
    APPEND LINES OF
           VALUE ttb_button(
                    ( butn_type = 3 )  " separator
                    ( function = 'EXPO' icon = icon_export
                      text = 'Export' quickinfo = 'export texts to xls'
                      disabled = ' ' )
           )
         TO e_object->mt_toolbar.
  ENDMETHOD.

  METHOD after_refresh.
    gr_report_view->change_subtotals( ).
  ENDMETHOD.

  METHOD export_excel.
    CONSTANTS:
      c_red_cell   TYPE lvc_istyle VALUE '7',
      c_green_cell TYPE lvc_istyle VALUE '6',
      c_subtotal   TYPE lvc_istyle VALUE '36',
      c_total      TYPE lvc_istyle VALUE '44'.

    CONSTANTS:
      c_header_count TYPE i VALUE 1.

    DATA :
      lr_grid_facade TYPE REF TO cl_salv_gui_grid_facade.

    DATA: lo_excel        TYPE REF TO zcl_excel,
          lo_excel_writer TYPE REF TO zif_excel_writer,
          lo_worksheet    TYPE REF TO zcl_excel_worksheet.
*          lo_column       TYPE REF TO zcl_excel_worksheet_columndime.

    DATA : lv_style TYPE zexcel_cell_style.

    DATA : lv_progress_percentage TYPE i VALUE 0.

    FIELD-SYMBOLS:
      <lt_data2> TYPE lvc_t_data,
      <lt_info2> TYPE lvc_t_info.

    TYPES:
      tt_data TYPE SORTED TABLE OF lvc_s_data WITH NON-UNIQUE KEY row_pos col_pos,
      tt_info TYPE SORTED TABLE OF lvc_s_info WITH NON-UNIQUE KEY col_pos.

    DATA:
      lt_data TYPE tt_data,
      lt_info TYPE tt_info.

*GET ALV DATA
**********************************************************************
    CREATE OBJECT lr_grid_facade
      EXPORTING
        o_grid = gr_report_view->gr_grid.


    CALL FUNCTION 'SAPGUI_PROGRESS_INDICATOR'
      EXPORTING
        percentage = lv_progress_percentage
        text       = |Extracting Alv ... %{ lv_progress_percentage }|.

    lr_grid_facade->if_salv_gui_grid_lvc_data~get_all_data(
      EXPORTING
        gui_type      = cl_salv_gru_view_grid=>c_gui_type-windows
        view          = if_salv_c_function=>view_excel
      IMPORTING
        table_lines   = DATA(lt_table_lines)
        rt_data       = DATA(lr_data)
        rt_info       = DATA(lr_info)
        rt_idpo       = DATA(lr_idpo)
        rt_poid       = DATA(lr_poid)
        rt_roid       = DATA(lr_roid)
        t_start_index = DATA(lr_start_index)
    ).

    ASSIGN lr_data->* TO <lt_data2>.
    ASSIGN lr_info->* TO <lt_info2>.

    INSERT LINES OF <lt_data2> INTO TABLE lt_data.
    INSERT LINES OF <lt_info2> INTO TABLE lt_info.


    lv_progress_percentage = 100.
    CALL FUNCTION 'SAPGUI_PROGRESS_INDICATOR'
      EXPORTING
        percentage = lv_progress_percentage
        text       = |Extracting Alv ... %{ lv_progress_percentage }|.

*GET ACTIVE SHEET
**********************************************************************
    CREATE OBJECT lo_excel.
    CREATE OBJECT lo_excel_writer TYPE zcl_excel_writer_2007.
    TRY.
        lo_worksheet = lo_excel->get_active_worksheet( ).
        lo_worksheet->set_title( ip_title = CONV #( sy-title ) ).
      CATCH zcx_excel.
        EXIT.
    ENDTRY.

*SET STYLE
**********************************************************************
    "header line
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA(lo_title_style) = lo_excel->add_new_style( ).
    lo_title_style->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
    lo_title_style->fill->fgcolor-rgb  = zcl_excel_style_color=>c_gray.
    lo_title_style->font->bold = abap_true.
    lo_title_style->font->size = 10.
    lo_title_style->alignment->wraptext = 'X'.
    lo_title_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_title_style->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
    CREATE OBJECT lo_title_style->borders->allborders.
    lo_title_style->borders->allborders->border_style = zcl_excel_style_border=>c_border_thin.
    DATA(lv_style_title) = lo_title_style->get_guid( ).

    "total line style
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA(lo_total_style) = lo_excel->add_new_style( ).
    lo_total_style->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
    lo_total_style->fill->fgcolor-rgb  = zcl_excel_style_color=>c_darkyellow.
    lo_total_style->font->bold = abap_true.
    lo_total_style->font->size = 10.
    lo_total_style->alignment->wraptext = 'X'.
    lo_total_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_total_style->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
    CREATE OBJECT lo_total_style->borders->allborders.
    lo_total_style->borders->allborders->border_style = zcl_excel_style_border=>c_border_thin.
    DATA(lv_style_total) = lo_total_style->get_guid( ).

    "subtotal style - sides border
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA(lo_subtotal_style1) = lo_excel->add_new_style( ).
    lo_subtotal_style1->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
    lo_subtotal_style1->fill->fgcolor-rgb  = zcl_excel_style_color=>c_yellow.
    lo_subtotal_style1->font->bold = abap_true.
    lo_subtotal_style1->font->size = 10.
    lo_subtotal_style1->alignment->wraptext = 'X'.
    lo_subtotal_style1->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_subtotal_style1->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
    CREATE OBJECT lo_subtotal_style1->borders->left.
    CREATE OBJECT lo_subtotal_style1->borders->right.
    lo_subtotal_style1->borders->left->border_style = zcl_excel_style_border=>c_border_thin.
    lo_subtotal_style1->borders->right->border_style = zcl_excel_style_border=>c_border_thin.
    DATA(lv_style_subtotal1) = lo_subtotal_style1->get_guid( ).

    "subtotal style - down border
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA(lo_subtotal_style2) = lo_excel->add_new_style( ).
    lo_subtotal_style2->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
    lo_subtotal_style2->fill->fgcolor-rgb  = zcl_excel_style_color=>c_yellow.
    lo_subtotal_style2->font->bold = abap_true.
    lo_subtotal_style2->font->size = 10.
    lo_subtotal_style2->alignment->wraptext = 'X'.
    lo_subtotal_style2->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_subtotal_style2->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
    CREATE OBJECT lo_subtotal_style2->borders->down.
    lo_subtotal_style2->borders->down->border_style = zcl_excel_style_border=>c_border_thin.
    DATA(lv_style_subtotal2) = lo_subtotal_style2->get_guid( ).

    "subtotal style - top down border
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA(lo_subtotal_style3) = lo_excel->add_new_style( ).
    lo_subtotal_style3->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
    lo_subtotal_style3->fill->fgcolor-rgb  = zcl_excel_style_color=>c_yellow.
    lo_subtotal_style3->font->bold = abap_true.
    lo_subtotal_style3->font->size = 10.
    lo_subtotal_style3->alignment->wraptext = 'X'.
    lo_subtotal_style3->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_subtotal_style3->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
    CREATE OBJECT lo_subtotal_style3->borders->top.
    CREATE OBJECT lo_subtotal_style3->borders->down.
    lo_subtotal_style3->borders->top->border_style = zcl_excel_style_border=>c_border_thin.
    lo_subtotal_style3->borders->down->border_style = zcl_excel_style_border=>c_border_thin.
    DATA(lv_style_subtotal3) = lo_subtotal_style3->get_guid( ).

    "cell - green
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA(lo_green_style) = lo_excel->add_new_style( ).
    lo_green_style->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
    lo_green_style->fill->fgcolor-rgb  = zcl_excel_style_color=>c_green.
    lo_green_style->font->size = 10.
    lo_green_style->alignment->wraptext = 'X'.
    lo_green_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_green_style->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
    CREATE OBJECT lo_green_style->borders->allborders.
    lo_green_style->borders->allborders->border_style = zcl_excel_style_border=>c_border_thin.
    DATA(lv_style_green_cell) = lo_green_style->get_guid( ).

    "cell - red
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA(lo_red_style) = lo_excel->add_new_style( ).
    lo_red_style->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
    lo_red_style->fill->fgcolor-rgb  = zcl_excel_style_color=>c_red.
    lo_red_style->font->size = 10.
    lo_red_style->alignment->wraptext = 'X'.
    lo_red_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_red_style->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
    CREATE OBJECT lo_red_style->borders->allborders.
    lo_red_style->borders->allborders->border_style = zcl_excel_style_border=>c_border_thin.
    DATA(lv_style_red_cell) = lo_red_style->get_guid( ).

    "cell others
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA(lo_cell_style) = lo_excel->add_new_style( ).
    lo_cell_style->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
    lo_cell_style->fill->fgcolor-rgb  = zcl_excel_style_color=>c_white.
    lo_cell_style->font->size = 10.
    lo_cell_style->alignment->wraptext = 'X'.
    lo_cell_style->alignment->horizontal = zcl_excel_style_alignment=>c_horizontal_center.
    lo_cell_style->alignment->vertical = zcl_excel_style_alignment=>c_vertical_center.
    CREATE OBJECT lo_cell_style->borders->allborders.
    lo_cell_style->borders->allborders->border_style = zcl_excel_style_border=>c_border_thin.
    DATA(lv_style_cell) = lo_cell_style->get_guid( ).



*DATA TO EXCEL
**********************************************************************
    "sortingenli sütunları bulmak
    DATA:
      lv_colpos1 TYPE i,
      lv_colpos2 TYPE i,
      lv_colpos3 TYPE i,
      lv_colpos4 TYPE i,
      lv_colpos5 TYPE i,
      lv_colpos6 TYPE i,
      lv_colpos7 TYPE i,
      lv_colpos8 TYPE i.


    "header
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    LOOP AT lt_info ASSIGNING FIELD-SYMBOL(<lfs_info>).
      TRY.
          lo_worksheet->set_cell( ip_row = 1
                                  ip_column = <lfs_info>-col_pos
                                  ip_value = <lfs_info>-text_m
                                  ip_style = lv_style_title   ).
        CATCH zcx_excel.
      ENDTRY.

      CHECK <lfs_info>-merge IS NOT INITIAL.
      IF lv_colpos1 IS INITIAL.
        lv_colpos1 = <lfs_info>-col_pos.
        CONTINUE.
      ENDIF.
      IF lv_colpos2 IS INITIAL.
        lv_colpos2 = <lfs_info>-col_pos.
        CONTINUE.
      ENDIF.
      IF lv_colpos3 IS INITIAL.
        lv_colpos3 = <lfs_info>-col_pos.
        CONTINUE.
      ENDIF.
      IF lv_colpos4 IS INITIAL.
        lv_colpos4 = <lfs_info>-col_pos.
        CONTINUE.
      ENDIF.
      IF lv_colpos5 IS INITIAL.
        lv_colpos5 = <lfs_info>-col_pos.
        CONTINUE.
      ENDIF.
      IF lv_colpos6 IS INITIAL.
        lv_colpos6 = <lfs_info>-col_pos.
        CONTINUE.
      ENDIF.
      IF lv_colpos7 IS INITIAL.
        lv_colpos7 = <lfs_info>-col_pos.
        CONTINUE.
      ENDIF.
      IF lv_colpos8 IS INITIAL.
        lv_colpos8 = <lfs_info>-col_pos.
        CONTINUE.
      ENDIF.

    ENDLOOP.

    "data
    """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    DATA :
      lv_sub1beg TYPE i VALUE 2,
      lv_sub2beg TYPE i VALUE 2,
      lv_sub3beg TYPE i VALUE 2,
      lv_sub4beg TYPE i VALUE 2,
      lv_sub5beg TYPE i VALUE 2,
      lv_sub6beg TYPE i VALUE 2,
      lv_sub7beg TYPE i VALUE 2,
      lv_sub8beg TYPE i VALUE 2,
      lv_totbeg  TYPE i VALUE 2.

    DATA : lv_sub_count TYPE i VALUE 0.

    DATA : lv_row_from TYPE i,
           lv_row_to   TYPE i.

    DATA : lv_rowno TYPE zexcel_cell_row.

    DATA(lv_lines) = lines( lt_data ).
    READ TABLE lt_data ASSIGNING FIELD-SYMBOL(<lfs_data2>) INDEX lv_lines.
    DATA(lv_last_row) = <lfs_data2>-row_pos.

    LOOP AT lt_data ASSIGNING FIELD-SYMBOL(<lgfs_row>) WHERE col_pos GT 0 GROUP BY <lgfs_row>-row_pos .

      LOOP AT GROUP <lgfs_row> ASSIGNING FIELD-SYMBOL(<lfs_column>) WHERE col_pos GT 0.

        "style
        """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        CASE <lfs_column>-style.
          WHEN c_red_cell.
            lv_style = lv_style_red_cell.
          WHEN c_green_cell.
            lv_style = lv_style_green_cell.
          WHEN c_subtotal.

            READ TABLE lt_info ASSIGNING <lfs_info> WITH KEY col_pos = <lfs_column>-col_pos BINARY SEARCH.


            IF <lfs_column>-value IS INITIAL OR
                <lfs_info>-merge NE 'X'.
              lv_style = lv_style_subtotal3.
            ELSE.
              READ TABLE lt_data ASSIGNING FIELD-SYMBOL(<lfs_data>)
                  WITH KEY col_pos = <lfs_column>-col_pos
                           row_pos = <lfs_column>-row_pos + 1 BINARY SEARCH.
              IF <lfs_data>-value EQ <lfs_column>-value.
                lv_style = lv_style_subtotal1."sonraki değer değişmişse->sides border
              ELSE.
                lv_style = lv_style_subtotal2."sonraki değer değişmemişse->down border
              ENDIF.
            ENDIF.

          WHEN c_total.
            lv_style = lv_style_total.
          WHEN OTHERS.

            IF <lfs_column>-col_pos EQ lv_colpos1 OR
               <lfs_column>-col_pos EQ lv_colpos2 OR
               <lfs_column>-col_pos EQ lv_colpos3 OR
               <lfs_column>-col_pos EQ lv_colpos4 OR
               <lfs_column>-col_pos EQ lv_colpos5 OR
               <lfs_column>-col_pos EQ lv_colpos6 OR
               <lfs_column>-col_pos EQ lv_colpos7 OR
               <lfs_column>-col_pos EQ lv_colpos8.

              lv_style = lv_style_subtotal1.

            ELSE.

              lv_style = lv_style_cell.

            ENDIF.
        ENDCASE.
        "row no
        """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        lv_rowno = <lgfs_row>-row_pos + c_header_count.

        "value
        """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        CONDENSE <lfs_column>-value NO-GAPS.

        "set cell
        """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        TRY.
            lo_worksheet->set_cell( ip_row = lv_rowno
                                    ip_column = <lfs_column>-col_pos
                                    ip_value = <lfs_column>-value
                                    ip_abap_type = cl_abap_typedescr=>typekind_string
                                    ip_style = COND #( WHEN lv_style IS NOT INITIAL THEN lv_style )  ).
          CATCH zcx_excel.
        ENDTRY.

        CLEAR : lv_style.
      ENDLOOP.

      "grouping
      """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
      IF <lgfs_row>-style EQ c_subtotal.
        lv_sub_count = lv_sub_count + 1.

        CASE lv_sub_count.
          WHEN 1.
            lv_row_from = lv_sub1beg.
          WHEN 2.
            lv_row_from = lv_sub2beg.
          WHEN 3.
            lv_row_from = lv_sub3beg.
          WHEN 4.
            lv_row_from = lv_sub4beg.
          WHEN 5.
            lv_row_from = lv_sub5beg.
          WHEN 6.
            lv_row_from = lv_sub6beg.
          WHEN 7.
            lv_row_from = lv_sub7beg.
          WHEN 8.
            lv_row_from = lv_sub8beg.
          WHEN 9.
            lv_row_from = lv_totbeg.
        ENDCASE.

        lv_row_to = <lgfs_row>-row_pos .

        TRY.
            lo_worksheet->set_row_outline(
              EXPORTING
                iv_row_from  = lv_row_from
                iv_row_to    = lv_row_to
                iv_collapsed = abap_false
            ).
          CATCH zcx_excel.
        ENDTRY.

      ELSE.
        IF <lgfs_row>-style NE c_total.
          IF lv_sub_count NE 0.

            CASE lv_sub_count.
              WHEN 1.
                lv_sub1beg = <lgfs_row>-row_pos + c_header_count.
              WHEN 2.
                lv_sub1beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub2beg = <lgfs_row>-row_pos + c_header_count.
              WHEN 3.
                lv_sub1beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub2beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub3beg = <lgfs_row>-row_pos + c_header_count.
              WHEN 4.
                lv_sub1beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub2beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub3beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub4beg = <lgfs_row>-row_pos + c_header_count.
              WHEN 5.
                lv_sub1beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub2beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub3beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub4beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub5beg = <lgfs_row>-row_pos.
              WHEN 6.
                lv_sub1beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub2beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub3beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub4beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub5beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub6beg = <lgfs_row>-row_pos + c_header_count.
              WHEN 7.
                lv_sub1beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub2beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub3beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub4beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub5beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub6beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub7beg = <lgfs_row>-row_pos + c_header_count.
              WHEN 8.
                lv_sub1beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub2beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub3beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub4beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub5beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub6beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub7beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub8beg = <lgfs_row>-row_pos + c_header_count.
              WHEN 9 .
                lv_sub1beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub2beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub3beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub4beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub5beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub6beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub7beg = <lgfs_row>-row_pos + c_header_count.
                lv_sub8beg = <lgfs_row>-row_pos + c_header_count.
                lv_totbeg = <lgfs_row>-row_pos + c_header_count.
            ENDCASE.
          ENDIF.
          lv_sub_count = 0.
        ENDIF.
      ENDIF.
      CLEAR : lv_row_from,lv_row_to.


      lv_progress_percentage = <lgfs_row>-row_pos * 100 / lv_last_row.
      CALL FUNCTION 'SAPGUI_PROGRESS_INDICATOR'
        EXPORTING
          percentage = lv_progress_percentage
          text       = |ALV to Excel ... %{ lv_progress_percentage }|.
    ENDLOOP.


*EXCEL TO XSTRING
**********************************************************************
    TRY.
        DATA(lv_highest_column) = lo_worksheet->get_highest_column( ).
        DATA(lv_count) = CONV int4( '1' ).

        WHILE lv_count <= lv_highest_column.
          DATA(lv_col_alpha) = zcl_excel_common=>convert_column2alpha( ip_column = lv_count ).
          DATA(lo_column_dimension) = lo_worksheet->get_column_dimension( ip_column = lv_col_alpha ).
          lo_column_dimension->set_auto_size( ip_auto_size = abap_true ).

          lv_count = lv_count + 1.
        ENDWHILE.
      CATCH zcx_excel.
        EXIT.
    ENDTRY.

    lv_progress_percentage = 0.
    CALL FUNCTION 'SAPGUI_PROGRESS_INDICATOR'
      EXPORTING
        percentage = lv_progress_percentage
        text       = |Excel downloading ... %{ lv_progress_percentage }|.

    rv_xlsx = lo_excel_writer->write_file( lo_excel ).


    lv_progress_percentage = 100.
    CALL FUNCTION 'SAPGUI_PROGRESS_INDICATOR'
      EXPORTING
        percentage = lv_progress_percentage
        text       = |Excel downloading ... %{ lv_progress_percentage }|.

  ENDMETHOD.

ENDCLASS.

*__ Class View ______________________________________________________*
CLASS lcl_report_view IMPLEMENTATION.
  METHOD get_instance.
    IF mo_view IS NOT BOUND.
      CREATE OBJECT mo_view.
    ENDIF.
    ro_view = mo_view.
  ENDMETHOD.

  METHOD display_data.
    IF gr_grid IS NOT BOUND.
      me->build_fcat( EXPORTING i_str = c_str_h
                      CHANGING t_fcat = gt_fcat ).
      me->_set_sort( ).
      me->exclude_functions( ).
      me->display_alv( ).
      me->change_subtotals( ).
      me->refresh_alv( gr_grid ).

    ELSE.
      me->change_subtotals( ).
      me->refresh_alv( gr_grid ).
    ENDIF.
  ENDMETHOD.

  METHOD refresh_alv.
    ir_grid->refresh_table_display( is_stable = VALUE #( row = 'X' col = 'X' )
                                    i_soft_refresh = 'X' ).
  ENDMETHOD.

  METHOD display_alv.
    gr_cont = NEW #( container_name = 'MAIN' ).
    gr_grid = NEW #( i_parent = gr_cont ).


    gr_event_handler = NEW #( iv_event = 'HDR' ).
    SET HANDLER gr_event_handler->handle_user_command FOR gr_grid.
    SET HANDLER gr_event_handler->handle_toolbar FOR gr_grid.
*    SET HANDLER gr_event_handler->handle_print_top_of_list FOR gr_grid.
*    SET HANDLER gr_event_handler->after_refresh FOR gr_grid.


    gr_grid->set_table_for_first_display(
      EXPORTING
        is_variant                    = VALUE #( report = sy-repid
                                                 username = sy-uname
                                                 handle = 'MAIN' )
        i_save                        = 'A'
        is_layout                     = VALUE #( zebra = 'X'
                                                 cwidth_opt = 'A'
                                                 sel_mode = 'A'
                                                 ctab_fname = 'COLOR'
                                                )
        it_toolbar_excluding          = gs_toolbar_excluding
      CHANGING
        it_outtab                     = gr_main->mt_data
        it_fieldcatalog               = me->gt_fcat
        it_sort                       = me->gt_sort
*        it_filter                     =
      EXCEPTIONS
        invalid_parameter_combination = 1
        program_error                 = 2
        too_many_lines                = 3
        OTHERS                        = 4
    ).
    IF sy-subrc <> 0.
      MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
        WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
    ENDIF.


  ENDMETHOD .

  METHOD exclude_functions.
    gs_toolbar_excluding = VALUE #( ( cl_gui_alv_grid=>mc_fc_print )
                                    ( cl_gui_alv_grid=>mc_fc_loc_append_row )
                                    ( cl_gui_alv_grid=>mc_fc_loc_insert_row )
                                    ( cl_gui_alv_grid=>mc_fc_loc_copy )
                                    ( cl_gui_alv_grid=>mc_fc_loc_cut )
                                    ( cl_gui_alv_grid=>mc_fc_loc_paste )
                                    ( cl_gui_alv_grid=>mc_fc_loc_paste_new_row )
                                    ( cl_gui_alv_grid=>mc_fc_loc_copy_row )
                                    ( cl_gui_alv_grid=>mc_fc_loc_delete_row )
                                    ( cl_gui_alv_grid=>mc_fc_loc_undo )
                                    ( cl_gui_alv_grid=>mc_fc_url_copy_to_clipboard )
*                                    ( cl_gui_alv_grid=>mc_fc_sort_dsc )
*                                    ( cl_gui_alv_grid=>mc_fc_refresh )
*                                    ( cl_gui_alv_grid=>mc_fc_check )
*                                    ( cl_gui_alv_grid=>mc_fc_loc_move_row )
*                                    ( cl_gui_alv_grid=>mc_mb_sum )
*                                    ( cl_gui_alv_grid=>mc_mb_subtot )
*                                    ( cl_gui_alv_grid=>mc_fc_graph )
*                                    ( cl_gui_alv_grid=>mc_fc_info )
*                                    ( cl_gui_alv_grid=>mc_fc_print_back )
*                                    ( cl_gui_alv_grid=>mc_fc_filter )
*                                    ( cl_gui_alv_grid=>mc_fc_find_more )
*                                    ( cl_gui_alv_grid=>mc_fc_find )
*                                    ( cl_gui_alv_grid=>mc_mb_export )
*                                    ( cl_gui_alv_grid=>mc_mb_variant )
*                                    ( cl_gui_alv_grid=>mc_fc_detail )
*                                    ( cl_gui_alv_grid=>mc_mb_view )
                                   ).
  ENDMETHOD.


  METHOD build_fcat.
    DEFINE m_set_coltext.
      &1-coltext   = &2.
      &1-scrtext_s = &2.
      &1-scrtext_m = &2.
      &1-scrtext_l = &2.
    END-OF-DEFINITION.

    REFRESH t_fcat.
    CALL FUNCTION 'LVC_FIELDCATALOG_MERGE'
      EXPORTING
        i_structure_name       = i_str
      CHANGING
        ct_fieldcat            = t_fcat
      EXCEPTIONS
        inconsistent_interface = 1
        program_error          = 2
        OTHERS                 = 3.
    IF sy-subrc NE 0.
      EXIT.
    ENDIF.

    LOOP AT t_fcat REFERENCE INTO DATA(gr_fcat).
      CASE gr_fcat->fieldname.
        WHEN 'KWMENG'.
          gr_fcat->do_sum = 'X'.
        WHEN 'NETWR'.
          gr_fcat->do_sum = 'X'.
        WHEN 'TEYIT_TTR'.
          gr_fcat->do_sum = 'X'.
        WHEN 'FIIL_SEVK_ADT'.
          gr_fcat->do_sum = 'X'.
        WHEN 'FIIL_SEVK_TTR'.
          gr_fcat->do_sum = 'X'.
        WHEN 'ADETSEL'.
          gr_fcat->do_sum = 'X'.
        WHEN 'TUTARSAL'.
          gr_fcat->do_sum = 'X'.
        WHEN 'CONFDQTYAFTERRUNINBASEUNIT'.
          gr_fcat->do_sum = 'X'.
        WHEN 'ABOPRUNUUID'.
          gr_fcat->tech = abap_true.
      ENDCASE.
    ENDLOOP.

  ENDMETHOD.        "build_fcat


  METHOD change_subtotals.
    gr_grid->get_subtotals(
      IMPORTING
        ep_collect00   = DATA(lt_00)
        ep_collect01   = DATA(lt_01)
        ep_collect02   = DATA(lt_02)
        ep_collect03   = DATA(lt_03)
    ).

    ASSIGN lt_00->* TO FIELD-SYMBOL(<t00>).
    <t00> = CORRESPONDING #( gr_main->gt_collect00 ).

    ASSIGN lt_01->* TO FIELD-SYMBOL(<t01>).
    <t01> = CORRESPONDING #( gr_main->gt_collect01 ).

    ASSIGN lt_02->* TO FIELD-SYMBOL(<t02>).
    <t02> = CORRESPONDING #( gr_main->gt_collect02 ).

    ASSIGN lt_03->* TO FIELD-SYMBOL(<t03>).
    <t03> = CORRESPONDING #( gr_main->gt_collect03 ).

  ENDMETHOD.


  METHOD _set_sort.
    DATA : ls_sort LIKE LINE OF gt_sort.

    ls_sort-fieldname = 'KVGR2'.
    ls_sort-up = 'X'.
    ls_sort-spos = '1'.
    ls_sort-subtot = 'X'.
    APPEND ls_sort TO gt_sort.

    ls_sort-fieldname = 'VBELN'.
    ls_sort-up = 'X'.
    ls_sort-spos = '2'.
    ls_sort-subtot = 'X'.
    APPEND ls_sort TO gt_sort.

    CLEAR : ls_sort.
    ls_sort-fieldname = 'POSNR'.
    ls_sort-up = 'X'.
    ls_sort-spos = '3'.
    ls_sort-subtot = 'X'.
    APPEND ls_sort TO gt_sort.

  ENDMETHOD.

ENDCLASS.
