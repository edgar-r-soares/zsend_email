CLASS zsend_email DEFINITION
  PUBLIC
  FINAL
  CREATE PUBLIC .

  PUBLIC SECTION.
    TYPES: BEGIN OF t_styles,
             header_up   TYPE zexcel_cell_style,
             header_edit TYPE zexcel_cell_style,
             header_down TYPE zexcel_cell_style,
             normal      TYPE zexcel_cell_style,
             edit        TYPE zexcel_cell_style,
             red         TYPE zexcel_cell_style,
             light_red   TYPE zexcel_cell_style,
             green       TYPE zexcel_cell_style,
             yellow      TYPE zexcel_cell_style,
             light_blue  TYPE zexcel_cell_style,
           END OF t_styles.

    METHODS add_table_to_excel
      IMPORTING
        !i_sheet_title TYPE zexcel_sheet_title OPTIONAL .
    METHODS finalize_excel_file
      EXPORTING
        !ot_rawdata  TYPE solix_tab
        !o_bytecount TYPE i .
    METHODS constructor
      IMPORTING
        !ir_salv  TYPE REF TO cl_salv_table OPTIONAL
        !ir_table TYPE REF TO data OPTIONAL .
    METHODS add_table_to_html
      IMPORTING
        VALUE(i_preview) TYPE sy-tabix OPTIONAL
      RETURNING
        VALUE(ot_html)   TYPE bcsy_text .
    CLASS-METHODS download_excel_file
      IMPORTING
        !ir_salv                  TYPE REF TO cl_salv_table
        !it_table                 TYPE ANY TABLE
        VALUE(i_sheet_title)      TYPE zexcel_sheet_title OPTIONAL
        VALUE(i_default_filename) TYPE string OPTIONAL .
    CLASS-METHODS salv_table
      IMPORTING
        !ir_salv             TYPE REF TO cl_salv_table
        !it_table            TYPE ANY TABLE
        VALUE(i_sender)      TYPE adr6-smtp_addr
        VALUE(i_receiver)    TYPE adr6-smtp_addr
        VALUE(i_mail_title)  TYPE so_obj_des OPTIONAL
        VALUE(i_attach)      TYPE boolean OPTIONAL
        VALUE(i_filename)    TYPE sood-objdes OPTIONAL
        !it_header           TYPE bcsy_text OPTIONAL
        !it_footer           TYPE bcsy_text OPTIONAL
        VALUE(i_sheet_title) TYPE zexcel_sheet_title OPTIONAL
        VALUE(i_preview)     TYPE sy-tabix DEFAULT 5 .
    CLASS-METHODS simple_message
      IMPORTING
        VALUE(i_sender)     TYPE adr6-smtp_addr
        VALUE(i_receiver)   TYPE adr6-smtp_addr
        VALUE(i_mail_title) TYPE so_obj_des OPTIONAL
        !it_message         TYPE bcsy_text OPTIONAL .
protected section.
private section.

  data R_EXCEL type ref to ZCL_EXCEL .
  data R_TABLE type ref to DATA .
  data R_SALV type ref to CL_SALV_TABLE .
  data S_STYLES type T_STYLES .

  methods GET_CELL_STYLE
    importing
      !I_NAME type STRING
      !IR_LINE type ref to DATA
    returning
      value(O_STYLE) type ZEXCEL_CELL_STYLE .
  methods GET_TD_CLASS
    importing
      !I_NAME type STRING
      !IR_LINE type ref to DATA
    returning
      value(O_ID) type STRING .
  methods INIT_EXCEL_STYLES .
  methods INIT_HTML_STYLES
    returning
      value(OT_HTML) type BCSY_TEXT .
ENDCLASS.



CLASS ZSEND_EMAIL IMPLEMENTATION.


  method ADD_TABLE_TO_EXCEL.
    data: lo_column      type ref to cl_salv_column.
  data: lo_columns     type ref to cl_salv_columns_table.
  data: lt_col type salv_t_column_ref.
  field-symbols: <ls_col> type line of salv_t_column_ref.
  data: l_s_text type scrtext_s.
  data: l_col type i.
  field-symbols: <lt_table> type standard table,
                 <ls_line> type any,
                 <l_field> type any.
  data: lo_type  type ref to cl_abap_typedescr,
        lo_element  type ref to cl_abap_elemdescr,
        l_name type string.
  data: l_td type string.
  data: l_field(255).
  data: lr_line type ref to data.

  data: lo_worksheet type ref to zcl_excel_worksheet.
  data: l_row type i.
  data: l_style type zexcel_cell_style.
  data: l_color_column type lvc_fname.
  data: lr_descr type ref to cl_abap_structdescr.
  field-symbols: <ls_comp> type abap_compdescr.


  lo_worksheet = me->r_excel->get_active_worksheet( ).
  lo_worksheet->set_title( ip_title = i_sheet_title ).


  assign me->r_table->* to <lt_table>.
  read table <lt_table> assigning <ls_line> index 1.
  if sy-subrc eq 0.

    lo_columns = r_salv->get_columns( ).
    l_color_column = lo_columns->get_color_column( ).
    lt_col = lo_columns->get( ).
    lr_descr ?= cl_abap_structdescr=>describe_by_data( <ls_line> ).

    l_col = 1.
    loop at lr_descr->components assigning <ls_comp> where name ne l_color_column.
      assign component <ls_comp>-name of structure <ls_line> to <l_field>.
      if sy-subrc eq 0.
        read table lt_col assigning <ls_col> with key columnname = <ls_comp>-name.
        l_s_text = <ls_col>-r_column->get_short_text( ).
        lo_worksheet->set_cell( ip_column = l_col ip_row = 1 ip_value = l_s_text ip_style = me->s_styles-header_up ).
        add 1 to l_col.
      endif.
    endloop.

    l_row = 1.
    loop at <lt_table> assigning <ls_line>.
      add 1 to l_row.
      l_col = 1.

      loop at lr_descr->components assigning <ls_comp> where name ne l_color_column.
        assign component <ls_comp>-name of structure <ls_line> to <l_field>.
        if sy-subrc eq 0.
          get reference of <ls_line> into lr_line.
          l_name = <ls_comp>-name.
          l_style = me->get_cell_style( i_name = l_name ir_line = lr_line ).
          lo_worksheet->set_cell( ip_column = l_col ip_row = l_row ip_value = <l_field> ip_style = l_style ).
          add 1 to l_col.
        endif.
      endloop.
    endloop.
  endif.

  endmethod.


method add_table_to_html.
  data: lo_column      type ref to cl_salv_column.
  data: lo_columns     type ref to cl_salv_columns_table.
  data: lt_col type salv_t_column_ref.
  field-symbols: <ls_col> type line of salv_t_column_ref.
  data: l_s_text type scrtext_s.
  field-symbols: <lt_table> type standard table,
                 <ls_line> type any,
                 <l_field> type any.
  data: lo_element  type ref to cl_abap_elemdescr,
        l_name type string.
  data: l_td type string.
  data: l_field(255)." TYPE string.
  data: lr_line type ref to data.
  data: l_color_column type lvc_fname.
  data: lr_descr type ref to cl_abap_structdescr.
  field-symbols: <ls_comp> type abap_compdescr.
  data: l_line type sy-tabix.

  assign me->r_table->* to <lt_table>.
  read table <lt_table> assigning <ls_line> index 1.
  if sy-subrc eq 0.

    lo_columns = r_salv->get_columns( ).
    l_color_column = lo_columns->get_color_column( ).
    lt_col = lo_columns->get( ).
    lr_descr ?= cl_abap_structdescr=>describe_by_data( <ls_line> ).

    append '<p><table>' to ot_html.

    append '<tr>' to ot_html.

    loop at lr_descr->components assigning <ls_comp> where name ne l_color_column.
      assign component <ls_comp>-name of structure <ls_line> to <l_field>.
      if sy-subrc eq 0.
        read table lt_col assigning <ls_col> with key columnname = <ls_comp>-name.
        l_s_text = <ls_col>-r_column->get_short_text( ).
        append '<th class="Heading">' to ot_html.
        append l_s_text to ot_html.
        append '</th>' to ot_html.
      endif.
    endloop.

    append '</tr>' to ot_html.

    loop at <lt_table> assigning <ls_line>.
      add 1 to l_line.
      append '<tr>' to ot_html.
      loop at lr_descr->components assigning <ls_comp> where name ne l_color_column.
        assign component <ls_comp>-name of structure <ls_line> to <l_field>.
        if sy-subrc eq 0.
          get reference of <ls_line> into lr_line.
          l_name = <ls_comp>-name.
          l_td = me->get_td_class( i_name = l_name ir_line = lr_line ).
          append l_td to ot_html.
          write <l_field> to l_field.
          append l_field to ot_html.
          append '</td>' to ot_html.
        endif.
      endloop.

      append '</tr>' to ot_html.

      if i_preview is not initial and l_line ge i_preview.
        l_line = lines( <lt_table> ).
        if l_line ge i_preview.
          append '<tr>' to ot_html.
          loop at lr_descr->components assigning <ls_comp> where name ne l_color_column.
            assign component <ls_comp>-name of structure <ls_line> to <l_field>.
            if sy-subrc eq 0.
              get reference of <ls_line> into lr_line.
              l_name = <ls_comp>-name.
              l_td = me->get_td_class( i_name = l_name ir_line = lr_line ).
              append l_td to ot_html.
              write '...' to l_field.
              append l_field to ot_html.
              append '</td>' to ot_html.
            endif.
          endloop.
        endif.
        exit.
      endif.

    endloop.

    append '</table></p>' to ot_html.

  endif.
endmethod.


method constructor.
  data: ls_vari type disvariant,
        lt_fcat type lvc_t_fcat,
        lt_sort type lvc_t_sort,
        lt_filt type lvc_t_filt,
        ls_layo type lvc_s_layo.
  field-symbols: <lt_table> type standard table.
  data: lr_table type ref to data.
  data lt_kkblo_fieldcat type kkblo_t_fieldcat.
  data ls_kkblo_layout  type kkblo_layout.
  data lt_kkblo_filter  type kkblo_t_filter.
  data lt_kkblo_sort    type kkblo_t_sortinfo.
  field-symbols: <ls_k_filter> type line of kkblo_t_filter.
  field-symbols: <ls_k_sort> type line of kkblo_t_sortinfo.
  data: lo_struct   type ref to cl_abap_structdescr,
*        lo_element  TYPE REF TO cl_abap_elemdescr,
        lo_tab      type ref to cl_abap_tabledescr,
        lt_comp     type cl_abap_structdescr=>component_table.
  field-symbols: <ls_comp> type line of cl_abap_structdescr=>component_table.
  data: lr_range type ref to data.
  field-symbols: <ls_line> type any,
                 <l_field> type any.
  field-symbols: <lt_rtable> type standard table,
                 <ls_rline> type any,
                 <l_rfield> type any.
  data: lt_sorttab  type abap_sortorder_tab.
  field-symbols: <ls_sort> type abap_sortorder.
  field-symbols: <ls_line_out> type any,
                 <lt_table_out> type standard table.
  data: lr_table_out type ref to data.
  field-symbols: <ls_k_field> type line of kkblo_t_fieldcat.
  field-symbols: <ls_l_filter> type line of lvc_t_filt.
  field-symbols: <ls_l_sort> type line of lvc_t_sort.
  data: lr_layout type ref to cl_salv_layout.
  data: lr_columns type ref to cl_salv_columns_table.
  data: lr_aggregations type ref to cl_salv_aggregations.
  field-symbols: <it_table> type standard table.

  if ir_salv is not initial.
    r_salv = ir_salv.
    assign ir_table->* to <it_table>.

    create data lr_table like <it_table>.
    assign lr_table->* to <lt_table>.
    <lt_table> = <it_table>.

    lr_layout = ir_salv->get_layout( ).

    cl_salv_controller_metadata=>get_variant(
      exporting
        r_layout  = lr_layout
      changing
        s_variant = ls_vari ).

    lr_columns = ir_salv->get_columns( ).
    lr_aggregations = ir_salv->get_aggregations( ).

    data: l_color_column type lvc_fname.
    l_color_column = lr_columns->get_color_column( ).


*... get the column information
    lt_fcat = cl_salv_controller_metadata=>get_lvc_fieldcatalog(
      r_columns      = lr_columns
      r_aggregations = lr_aggregations ).
    .

*... get the layout information
    cl_salv_controller_metadata=>get_lvc_layout(
      exporting
        r_columns             = lr_columns
        r_aggregations        = lr_aggregations
      changing
        s_layout              = ls_layo ).

* the fieldcatalog is not complete yet!
    call function 'LVC_FIELDCAT_COMPLETE'
      exporting
        i_complete       = 'X'
        i_refresh_buffer = space
        i_buffer_active  = space
        is_layout        = ls_layo
        i_test           = '1'
        i_fcat_complete  = 'X'
      changing
        ct_fieldcat      = lt_fcat.

    if ls_vari is not initial.

      call function 'LVC_TRANSFER_TO_KKBLO'
        exporting
          it_fieldcat_lvc         = lt_fcat
        importing
          et_fieldcat_kkblo       = lt_kkblo_fieldcat
        exceptions
          it_data_missing         = 1
          it_fieldcat_lvc_missing = 2
          others                  = 3.

      call function 'LT_VARIANT_LOAD'
        exporting
          i_tabname           = '1'
          i_dialog            = ' '
          i_fcat_complete     = 'X'
        importing
          et_fieldcat         = lt_kkblo_fieldcat
          et_sort             = lt_kkblo_sort
          et_filter           = lt_kkblo_filter
        changing
          cs_layout           = ls_kkblo_layout
          ct_default_fieldcat = lt_kkblo_fieldcat
          cs_variant          = ls_vari
        exceptions
          wrong_input         = 1
          fc_not_complete     = 2
          not_found           = 3
          others              = 4.

      call function 'LVC_TRANSFER_FROM_KKBLO'
        exporting
          it_fieldcat_kkblo = lt_kkblo_fieldcat
          it_sort_kkblo     = lt_kkblo_sort
          it_filter_kkblo   = lt_kkblo_filter
          is_layout_kkblo   = ls_kkblo_layout
        importing
          et_fieldcat_lvc   = lt_fcat
          et_sort_lvc       = lt_sort
          et_filter_lvc     = lt_filt
          es_layout_lvc     = ls_layo
        tables
          it_data           = <lt_table>
        exceptions
          it_data_missing   = 1
          others            = 2.

      loop at lt_kkblo_filter assigning <ls_k_filter>.
        read table <lt_table> assigning <ls_line> index 1.
        append initial line to lt_comp assigning <ls_comp>.
        <ls_comp>-name = 'SIGN'.
        <ls_comp>-type ?= cl_abap_typedescr=>describe_by_name( 'RALDB_SIGN' ).

        append initial line to lt_comp assigning <ls_comp>.
        <ls_comp>-name = 'OPTION'.
        <ls_comp>-type ?= cl_abap_typedescr=>describe_by_name( 'RALDB_OPTI' ).


        assign component <ls_k_filter>-fieldname of structure <ls_line> to <l_field>.
        append initial line to lt_comp assigning <ls_comp>.
        <ls_comp>-name = 'LOW'.
        <ls_comp>-type ?= cl_abap_datadescr=>describe_by_data( <l_field> ).
        append initial line to lt_comp assigning <ls_comp>.
        <ls_comp>-name = 'HIGH'.
        <ls_comp>-type ?= cl_abap_datadescr=>describe_by_data( <l_field> ).

        lo_struct = cl_abap_structdescr=>create( lt_comp ).
        lo_tab = cl_abap_tabledescr=>create( p_line_type = lo_struct p_table_kind = cl_abap_tabledescr=>tablekind_std p_unique = abap_false ).
        create data lr_range type handle lo_tab.
        assign lr_range->* to <lt_rtable>.
        append initial line to <lt_rtable> assigning <ls_rline>.
        assign component 'SIGN' of structure <ls_rline> to <l_rfield>.
        <l_rfield> = <ls_k_filter>-sign0.
        assign component 'OPTION' of structure <ls_rline> to <l_rfield>.
        <l_rfield> = <ls_k_filter>-optio.
        if <ls_k_filter>-valuf_int is initial and <ls_k_filter>-valut_int is initial.
          assign component 'LOW' of structure <ls_rline> to <l_rfield>.
          <l_rfield> = <ls_k_filter>-valuf.
          assign component 'HIGH' of structure <ls_rline> to <l_rfield>.
          <l_rfield> = <ls_k_filter>-valut.
        else.
          assign component 'LOW' of structure <ls_rline> to <l_rfield>.
          <l_rfield> = <ls_k_filter>-valuf_int.
          assign component 'HIGH' of structure <ls_rline> to <l_rfield>.
          <l_rfield> = <ls_k_filter>-valut_int.
        endif.

        loop at <lt_table> assigning <ls_line>.
          assign component <ls_k_filter>-fieldname of structure <ls_line> to <l_field>.
          if <l_field> not in <lt_rtable>.
            delete <lt_table>.
          endif.
        endloop.

        clear lt_comp.
      endloop.

      loop at lt_kkblo_sort assigning <ls_k_sort>.
        append initial line to lt_sorttab assigning <ls_sort>.
        <ls_sort>-name = <ls_k_sort>-fieldname.
        <ls_sort>-descending = <ls_k_sort>-down.
      endloop.
      if sy-subrc eq 0.
        sort <lt_table> by (lt_sorttab).
      endif.


      read table <it_table> assigning <ls_line> index 1.
      if sy-subrc eq 0.
        clear lt_comp.
        sort lt_kkblo_fieldcat by col_pos.
        loop at lt_kkblo_fieldcat assigning <ls_k_field> where ( no_out is initial and tech is initial ).
          assign component <ls_k_field>-fieldname of structure <ls_line> to <l_field>.
          append initial line to lt_comp assigning <ls_comp>.
          <ls_comp>-name = <ls_k_field>-fieldname.
          <ls_comp>-type ?= cl_abap_datadescr=>describe_by_data( <l_field> ).
        endloop.
        if l_color_column is not initial.
          data: lt_col type lvc_t_scol.
          append initial line to lt_comp assigning <ls_comp>.
          <ls_comp>-name = l_color_column.
          <ls_comp>-type ?= cl_abap_typedescr=>describe_by_data( lt_col ).
        endif.
        if lt_comp[] is not initial.
          lo_struct = cl_abap_structdescr=>create( lt_comp ).
          lo_tab = cl_abap_tabledescr=>create( p_line_type = lo_struct p_table_kind = cl_abap_tabledescr=>tablekind_std p_unique = abap_false ).
          create data lr_table_out type handle lo_tab.
          assign lr_table_out->* to <lt_table_out>.

          loop at <lt_table> assigning <ls_line>.
            append initial line to <lt_table_out> assigning <ls_line_out>.
            move-corresponding <ls_line> to <ls_line_out>.
          endloop.
        endif.
      endif.


    else.
*  ... get the sort information
      data: lr_sorts type ref to cl_salv_sorts.
      call method ir_salv->get_sorts
        receiving
          value = lr_sorts.
      lt_sort = cl_salv_controller_metadata=>get_lvc_sort( lr_sorts ).

*  ... get the filter information
      data: lr_filters type ref to cl_salv_filters.
      call method ir_salv->get_filters
        receiving
          value = lr_filters.

      lt_filt = cl_salv_controller_metadata=>get_lvc_filter( lr_filters ).


      loop at lt_filt assigning <ls_l_filter>.
        read table <lt_table> assigning <ls_line> index 1.
        append initial line to lt_comp assigning <ls_comp>.
        <ls_comp>-name = 'SIGN'.
        <ls_comp>-type ?= cl_abap_typedescr=>describe_by_name( 'RALDB_SIGN' ).

        <ls_comp>-name = 'OPTION'.
        <ls_comp>-type ?= cl_abap_typedescr=>describe_by_name( 'RALDB_OPTI' ).
        append initial line to lt_comp assigning <ls_comp>.


        assign component <ls_l_filter>-fieldname of structure <ls_line> to <l_field>.
        append initial line to lt_comp assigning <ls_comp>.
        <ls_comp>-name = 'LOW'.
        <ls_comp>-type ?= cl_abap_datadescr=>describe_by_data( <l_field> ).
        append initial line to lt_comp assigning <ls_comp>.
        <ls_comp>-name = 'HIGH'.
        <ls_comp>-type ?= cl_abap_datadescr=>describe_by_data( <l_field> ).

        if lt_comp[] is not initial.
          lo_struct = cl_abap_structdescr=>create( lt_comp ).
          lo_tab = cl_abap_tabledescr=>create( p_line_type = lo_struct p_table_kind = cl_abap_tabledescr=>tablekind_std p_unique = abap_false ).
          create data lr_range type handle lo_tab.
          assign lr_range->* to <lt_rtable>.
          append initial line to <lt_rtable> assigning <ls_rline>.
          assign component 'SIGN' of structure <ls_rline> to <l_rfield>.
          <l_rfield> = <ls_l_filter>-sign.
          assign component 'OPTION' of structure <ls_rline> to <l_rfield>.
          <l_rfield> = <ls_l_filter>-option.
          assign component 'LOW' of structure <ls_rline> to <l_rfield>.
          <l_rfield> = <ls_l_filter>-valuf.
          assign component 'HIGH' of structure <ls_rline> to <l_rfield>.
          <l_rfield> = <ls_l_filter>-valut.

          loop at <lt_table> assigning <ls_line>.
            assign component <ls_l_filter>-fieldname of structure <ls_line> to <l_field>.
            if <l_field> not in <lt_rtable>.
              delete <lt_table>.
            endif.
          endloop.
        endif.
        clear lt_comp.
      endloop.

      if lt_comp[] is not initial.
        loop at lt_sort assigning <ls_l_sort>.
          append initial line to lt_sorttab assigning <ls_sort>.
          <ls_sort>-name = <ls_l_sort>-fieldname.
          <ls_sort>-descending = <ls_l_sort>-down.
        endloop.
        if sy-subrc eq 0.
          sort <lt_table> by (lt_sorttab).
        endif.


        assign <lt_table_out> to <lt_table>.
      endif.
    endif.

    if <lt_table_out> is assigned.
      create data me->r_table like <lt_table_out>.
      assign me->r_table->* to <lt_table>.
      <lt_table> = <lt_table_out>.
    endif.


    create object r_excel.

  endif.

endmethod.


method download_excel_file.
  data: xdata       type xstring,             " Will be used for sending as email
        t_rawdata   type solix_tab,           " Will be used for downloading or open directly
        bytecount   type i.                   " Will be used for downloading or open directly
  data: l_save_file_name type string.
  data: lr_table type ref to zsend_email.
  data: lir_table type ref to data,
        l_file type string,
        l_path type string,
        l_fullpath type string.

  call method cl_gui_frontend_services=>file_save_dialog
    exporting
*      window_title         =
      default_extension    = 'XLSX'
      default_file_name    = i_default_filename
*      with_encoding        =
*      file_filter          =
*      initial_directory    =
*      prompt_on_overwrite  = 'X'
    changing
      filename             = l_file
      path                 = l_path
      fullpath             = l_fullpath
    exceptions
      cntl_error           = 1
      error_no_gui         = 2
      not_supported_by_gui = 3
      others               = 4
          .
  if sy-subrc eq 0 and l_file is not initial.

    get reference of it_table into lir_table.
    create object lr_table
      exporting
        ir_salv  = ir_salv
        ir_table = lir_table.
    lr_table->init_excel_styles( ).
    lr_table->add_table_to_excel( i_sheet_title ).
    lr_table->finalize_excel_file( importing ot_rawdata = t_rawdata o_bytecount = bytecount ).

    cl_gui_frontend_services=>gui_download( exporting bin_filesize = bytecount
                                                      filename     = l_file
                                                      filetype     = 'BIN'
                                             changing data_tab     = t_rawdata ).
  endif.




endmethod.


  METHOD finalize_excel_file.
    DATA: cl_writer TYPE REF TO zif_excel_writer.
    DATA: xdata TYPE xstring.             " Will be used for sending as email

    CREATE OBJECT cl_writer TYPE zcl_excel_writer_2007.
    xdata = cl_writer->write_file( me->r_excel ).
    ot_rawdata = cl_bcs_convert=>xstring_to_solix( iv_xstring  = xdata ).
    o_bytecount = xstrlen( xdata ).
  ENDMETHOD.


method get_cell_style.
  data: lo_columns     type ref to cl_salv_columns_table,
        l_color_column type lvc_fname.
  field-symbols: <ls_line> type any,
                 <lt_color> type lvc_t_scol,
                 <ls_color> type lvc_s_scol.

  lo_columns = me->r_salv->get_columns( ).
  l_color_column = lo_columns->get_color_column( ).

  if l_color_column is not initial.
    assign ir_line->* to <ls_line>.
    assign component l_color_column of structure <ls_line> to <lt_color>.
    read table <lt_color> assigning <ls_color> with key fname = i_name.
    if sy-subrc ne 0.
      o_style = s_styles-normal.
    else.
      case <ls_color>-color-col.
        when 1.
          o_style = s_styles-light_blue.
        when 3.
          o_style = s_styles-yellow.
        when 5.
          o_style = s_styles-green.
        when 6.
          o_style = s_styles-red.
        when 7.
          o_style = s_styles-light_red.
        when others.
          o_style = s_styles-normal.
      endcase.
    endif.
  else.
    o_style = s_styles-normal.
  endif.

endmethod.


method get_td_class.
  data: lo_columns     type ref to cl_salv_columns_table,
        l_color_column type lvc_fname.
  field-symbols: <ls_line> type any,
                 <lt_color> type lvc_t_scol,
                 <ls_color> type lvc_s_scol.

  lo_columns = me->r_salv->get_columns( ).
  l_color_column = lo_columns->get_color_column( ).

  if l_color_column is not initial.
    assign ir_line->* to <ls_line>.
    assign component l_color_column of structure <ls_line> to <lt_color>.
    read table <lt_color> assigning <ls_color> with key fname = i_name.
    if sy-subrc ne 0.
      o_id = '<TD CLASS="normal">'.
    else.
      case <ls_color>-color-col.
        when 1.
          o_id = '<TD CLASS="light_blue">'.
        when 3.
          o_id = '<TD CLASS="yellow">'.
        when 5.
          o_id = '<TD CLASS="green">'.
        when 6.
          o_id = '<TD CLASS="red">'.
        when 7.
          o_id = '<TD CLASS="light_red">'.
        when others.
          o_id = '<TD CLASS="normal">'.
      endcase.
    endif.
  else.
    o_id = '<TD CLASS="normal">'.
  endif.

endmethod.


method init_excel_styles.

  data: lo_worksheet            type ref to zcl_excel_worksheet,
        lo_style_header_up      type ref to zcl_excel_style,
        lo_style_header_edit    type ref to zcl_excel_style,
        lo_style_header_down    type ref to zcl_excel_style,
        lo_style_normal         type ref to zcl_excel_style,
        lo_style_edit           type ref to zcl_excel_style,
        lo_style_red            type ref to zcl_excel_style,
        lo_style_light_red      type ref to zcl_excel_style,
        lo_style_green          type ref to zcl_excel_style,
        lo_style_yellow         type ref to zcl_excel_style,
        lo_style_light_blue     type ref to zcl_excel_style,
        lo_border_dark          type ref to zcl_excel_style_border.

  " Create border object
  create object lo_border_dark.
  lo_border_dark->border_color-rgb = zcl_excel_style_color=>c_black.
  lo_border_dark->border_style = zcl_excel_style_border=>c_border_thin.

  "Create style top header
  lo_style_header_up                         = me->r_excel->add_new_style( ).
  lo_style_header_up->borders->allborders    = lo_border_dark.
  lo_style_header_up->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_header_up->fill->fgcolor-rgb  = '0F243E'. "top zahara blue
  lo_style_header_up->font->bold   = abap_true.
  lo_style_header_up->font->color-rgb  = zcl_excel_style_color=>c_white.
  lo_style_header_up->font->name         = zcl_excel_style_font=>c_name_calibri.
  lo_style_header_up->font->size = 10.
*  lo_style_header_up->font->scheme       = zcl_excel_style_font=>c_scheme_major.
*  lo_style_header_up->protection->locked = zcl_excel_style_protection=>c_protection_locked.
  s_styles-header_up                    = lo_style_header_up->get_guid( ).


  "Create style top header edit column
  lo_style_header_edit                         = me->r_excel->add_new_style( ).
  lo_style_header_edit->borders->allborders    = lo_border_dark.
  lo_style_header_edit->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_header_edit->fill->fgcolor-rgb  = '0F243E'. "top zahara blue
  lo_style_header_edit->font->bold        = abap_true.
  lo_style_header_edit->font->color-rgb  = zcl_excel_style_color=>c_red.
  lo_style_header_edit->font->name         = zcl_excel_style_font=>c_name_calibri.
  lo_style_header_edit->font->size = 10.
*  lo_style_header_edit->font->scheme       = zcl_excel_style_font=>c_scheme_major.
*  lo_style_header_edit->protection->locked = zcl_excel_style_protection=>c_protection_locked.
  s_styles-header_edit                    = lo_style_header_edit->get_guid( ).

  "Create style second line header
  lo_style_header_down                         = me->r_excel->add_new_style( ).
  lo_style_header_down->borders->allborders    = lo_border_dark.
  lo_style_header_down->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_header_down->fill->fgcolor-rgb  = 'ADD1A5'. "small header zahara green
  lo_style_header_down->font->name         = zcl_excel_style_font=>c_name_calibri.
  lo_style_header_down->font->size = 10.
*  lo_style_header_down->protection->locked = zcl_excel_style_protection=>c_protection_locked.
  s_styles-header_down                    = lo_style_header_down->get_guid( ).

  "Create style table cells
  lo_style_normal                         = me->r_excel->add_new_style( ).
  lo_style_normal->borders->allborders    = lo_border_dark.
  lo_style_normal->font->name             = zcl_excel_style_font=>c_name_calibri.
  lo_style_normal->font->size             = 8.
*  lo_style_normal->protection->locked = zcl_excel_style_protection=>c_protection_locked.
  s_styles-normal                         = lo_style_normal->get_guid( ).

  lo_style_red                         = me->r_excel->add_new_style( ).
  lo_style_red->borders->allborders    = lo_border_dark.
  lo_style_red->font->name             = zcl_excel_style_font=>c_name_calibri.
  lo_style_red->font->size             = 8.
  lo_style_red->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_red->fill->fgcolor-rgb  = 'FF3333'. "red
  s_styles-red                         = lo_style_red->get_guid( ).

  lo_style_light_red                         = me->r_excel->add_new_style( ).
  lo_style_light_red->borders->allborders    = lo_border_dark.
  lo_style_light_red->font->name             = zcl_excel_style_font=>c_name_calibri.
  lo_style_light_red->font->size             = 8.
  lo_style_light_red->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_light_red->fill->fgcolor-rgb  = 'FFCCCC'. "light_red
  s_styles-light_red                         = lo_style_light_red->get_guid( ).

  lo_style_green                         = me->r_excel->add_new_style( ).
  lo_style_green->borders->allborders    = lo_border_dark.
  lo_style_green->font->name             = zcl_excel_style_font=>c_name_calibri.
  lo_style_green->font->size             = 8.
  lo_style_green->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_green->fill->fgcolor-rgb  = '00CC44'. "green
  s_styles-green                         = lo_style_green->get_guid( ).

  lo_style_yellow                         = me->r_excel->add_new_style( ).
  lo_style_yellow->borders->allborders    = lo_border_dark.
  lo_style_yellow->font->name             = zcl_excel_style_font=>c_name_calibri.
  lo_style_yellow->font->size             = 8.
  lo_style_yellow->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_yellow->fill->fgcolor-rgb  = 'FFFF00'. "yellow
  s_styles-yellow                         = lo_style_yellow->get_guid( ).

  lo_style_light_blue                         = me->r_excel->add_new_style( ).
  lo_style_light_blue->borders->allborders    = lo_border_dark.
  lo_style_light_blue->font->name             = zcl_excel_style_font=>c_name_calibri.
  lo_style_light_blue->font->size             = 8.
  lo_style_light_blue->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_light_blue->fill->fgcolor-rgb  = 'DCE6F1'. "light blue
  s_styles-light_blue                         = lo_style_light_blue->get_guid( ).

  "Create style edit cells
  lo_style_edit                         = me->r_excel->add_new_style( ).
  lo_style_edit->borders->allborders    = lo_border_dark.
  lo_style_edit->font->name             = zcl_excel_style_font=>c_name_calibri.
  lo_style_edit->font->size             = 8.
  lo_style_edit->fill->filltype     = zcl_excel_style_fill=>c_fill_solid.
  lo_style_edit->fill->fgcolor-rgb  = 'FFFF66'. "top zahara blue
  lo_style_edit->number_format->format_code = zcl_excel_style_number_format=>c_format_text."c_format_number.
  lo_style_edit->protection->locked = zcl_excel_style_protection=>c_protection_unlocked.
  s_styles-edit                        = lo_style_edit->get_guid( ).

endmethod.


method init_html_styles.

  append '<style type="text/css">' to ot_html.
  append '<!--' to ot_html.
  append 'p{font-family: verdana; font-size:x-small}' to ot_html.
  append 'table{background-color:#FFF;border-collapse:collapse; font-family: verdana; font-size:xx-small}' to ot_html.
  append 'th.Heading{background-color:#0F243E;text-align:left;color:white;border:1px solidblack;padding:3px;font-family: verdana; font-size:xx-small}' to ot_html. "text-align:center;
  append 'td.red{background-color:#FF3333;border:1px solid black;padding:3px;font-family: verdana; font-size:xx-small}' to ot_html.
  append 'td.light_red{background-color:#FFCCCC;border:1px solid black;padding:3px;font-family: verdana; font-size:xx-small}' to ot_html.
  append 'td.green{background-color:#00CC44;border:1px solid black;padding:3px;font-family: verdana; font-size:xx-small}' to ot_html.
  append 'td.yellow{background-color:#FFFF00;border:1px solid black;padding:3px;font-family: verdana; font-size:xx-small}' to ot_html.
  append 'td.light_blue{background-color:#DCE6F1;border:1px solid black;padding:3px;font-family: verdana; font-size:xx-small}' to ot_html.
  append 'td.normal{background-color:#FFF;border:1px solid black;padding:3px;font-family: verdana; font-size:xx-small}' to ot_html.
  append '-->' to ot_html.
  append '</style>' to ot_html.

endmethod.


method salv_table.
  data: cl_writer type ref to zif_excel_writer.
  data: xdata       type xstring,             " Will be used for sending as email
        t_rawdata   type solix_tab,           " Will be used for downloading or open directly
        bytecount   type i.                   " Will be used for downloading or open directly
  data: l_save_file_name type string.
* Needed to send emails
  data: bcs_exception           type ref to cx_bcs,
        errortext               type string,
        cl_send_request         type ref to cl_bcs,
        cl_document             type ref to cl_document_bcs,
        cl_recipient            type ref to if_recipient_bcs,
        cl_sender               type ref to cl_cam_address_bcs,
        t_attachment_header     type soli_tab,
        wa_attachment_header    like line of t_attachment_header,
        attachment_subject      type sood-objdes,
        sood_bytecount          type sood-objlen,
        send_to                 type adr6-smtp_addr,
        sent                    type os_boolean.
  data: lt_html          type bcsy_text,
        lt_block         type bcsy_text,
        ls_html          type line of bcsy_text.
  field-symbols <ls_html> type line of bcsy_text.
  data: lr_table type ref to zsend_email.
  data: lir_table type ref to data.
  data: l_attach type boolean.
  data: l_bytes type f.

  get reference of it_table into lir_table.
  create object lr_table
    exporting
      ir_salv  = ir_salv
      ir_table = lir_table.

  append '<head>' to lt_html.
  lt_block = lr_table->init_html_styles( ).
  append lines of lt_block to lt_html.
  append '</head>' to lt_html.
  append '<body>' to lt_html.
  loop at it_header assigning <ls_html>.
    append '<p>' to lt_html.
    append <ls_html> to lt_html.
    append '</p>' to lt_html.
  endloop.

  describe table it_table.
  l_bytes = lines( it_table ) * sy-tleng.
  if i_attach is not initial and l_bytes gt 0.
    l_attach = i_attach.
  else.
    if l_bytes gt 6000000. "6MB, guessing email size, if too big then force attachment
      l_attach = abap_true.
    endif.
  endif.

  if l_attach is not initial.
    lr_table->init_excel_styles( ).
    lr_table->add_table_to_excel( i_sheet_title ).
    lr_table->finalize_excel_file( importing ot_rawdata = t_rawdata o_bytecount = bytecount ).

    lt_block = lr_table->add_table_to_html( i_preview ).
    append lines of lt_block to lt_html.

  else.
    if l_bytes gt 0.
      lt_block = lr_table->add_table_to_html( ).
      append lines of lt_block to lt_html.
    endif.
  endif.

  loop at it_footer assigning <ls_html>.
    append '<p>' to lt_html.
    append <ls_html> to lt_html.
    append '</p>' to lt_html.
  endloop.
  append '</body>' to lt_html.


  try.
* Create send request
      cl_send_request = cl_bcs=>create_persistent( ).

      data: lr_sender type ref to if_sender_bcs.
      lr_sender    = cl_cam_address_bcs=>create_internet_address( i_sender ).
      cl_send_request->set_sender( lr_sender ).
* Create new document with mailtitle and mailtextg
      cl_document = cl_document_bcs=>create_document( i_type    = 'HTM' "'RAW' "#EC NOTEXT
                                                      i_text    = lt_html "t_mailtext
                                                      i_subject = i_mail_title ).

      if l_attach is not initial.
* Add attachment to document
* since the new excelfiles have an 4-character extension .xlsx but the attachment-type only holds 3 charactes .xls,
* we have to specify the real filename via attachment header
* Use attachment_type xls to have SAP display attachment with the excel-icon
        attachment_subject  = i_filename.
        concatenate '&SO_FILENAME=' attachment_subject into wa_attachment_header.
        append wa_attachment_header to t_attachment_header.
* Attachment
        sood_bytecount = bytecount.  " next method expects sood_bytecount instead of any positive integer *sigh*
        cl_document->add_attachment(  i_attachment_type    = 'XLS' "#EC NOTEXT
                                      i_attachment_subject = attachment_subject
                                      i_attachment_size    = sood_bytecount
                                      i_att_content_hex    = t_rawdata
                                      i_attachment_header  = t_attachment_header ).
      endif.

* add document to send request
      cl_send_request->set_document( cl_document ).

* add recipient
      send_to = i_receiver.
      cl_recipient = cl_cam_address_bcs=>create_internet_address( send_to ).
      cl_send_request->add_recipient( cl_recipient ).

* Und abschicken
      sent = cl_send_request->send( i_with_error_screen = 'X' ).

*please note, there is a commit here
      commit work.

      if sent is initial.
        message i500(sbcoms) with i_receiver.
      else.
        message s022(so).
*        MESSAGE 'Document ready to be sent - Check SOST or SCOT' TYPE 'I'.
      endif.
    catch cx_bcs into bcs_exception.
      errortext = bcs_exception->if_message~get_text( ).
      message errortext type 'I'.
  endtry.


endmethod.


method simple_message.
* Needed to send emails
  data: bcs_exception           type ref to cx_bcs,
        errortext               type string,
        cl_send_request         type ref to cl_bcs,
        cl_document             type ref to cl_document_bcs,
        cl_recipient            type ref to if_recipient_bcs,
        cl_sender               type ref to cl_cam_address_bcs,
        sood_bytecount          type sood-objlen,
        send_to                 type adr6-smtp_addr,
        sent                    type os_boolean.
  data: lt_html          type bcsy_text,
        lt_block         type bcsy_text,
        ls_html          type line of bcsy_text.
  field-symbols <ls_html> type line of bcsy_text.
  data: lr_table type ref to zsend_email.
  data: lir_table type ref to data.

  create object lr_table.

  append '<head>' to lt_html.
  lt_block = lr_table->init_html_styles( ).
  append lines of lt_block to lt_html.
  append '</head>' to lt_html.
  append '<body>' to lt_html.
  loop at it_message assigning <ls_html>.
    append '<p>' to lt_html.
    append <ls_html> to lt_html.
    append '</p>' to lt_html.
  endloop.

  try.
* Create send request
      cl_send_request = cl_bcs=>create_persistent( ).

      data: lr_sender type ref to if_sender_bcs.
      lr_sender    = cl_cam_address_bcs=>create_internet_address( i_sender ).
      cl_send_request->set_sender( lr_sender ).
* Create new document with mailtitle and mailtextg
      cl_document = cl_document_bcs=>create_document( i_type    = 'HTM' "'RAW' "#EC NOTEXT
                                                      i_text    = lt_html "t_mailtext
                                                      i_subject = i_mail_title ).

* add document to send request
      cl_send_request->set_document( cl_document ).

* add recipient
      send_to = i_receiver.
      cl_recipient = cl_cam_address_bcs=>create_internet_address( send_to ).
      cl_send_request->add_recipient( cl_recipient ).

* Und abschicken
      sent = cl_send_request->send( i_with_error_screen = 'X' ).

*please note, there is a commit here
      commit work.

      if sent is initial.
        message i500(sbcoms) with i_receiver.
      else.
        message s022(so).
*        MESSAGE 'Document ready to be sent - Check SOST or SCOT' TYPE 'I'.
      endif.
    catch cx_bcs into bcs_exception.
      errortext = bcs_exception->if_message~get_text( ).
      message errortext type 'I'.
  endtry.


endmethod.
ENDCLASS.
