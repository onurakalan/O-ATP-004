TABLES: vbak, vbap, ekko, ekpo, mara, knvv, likp,caboprunreqconfn.

SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-h01.
  SELECT-OPTIONS: s_kunnr  FOR vbak-kunnr,
                  s_matnr  FOR vbap-matnr,
                  s_vbeln  FOR vbak-vbeln,
                  s_posnr  FOR caboprunreqconfn-atprelevantdocumentitem,
                  s_vkorg  FOR vbak-vkorg,
                  s_vtweg  FOR vbak-vtweg,
                  s_werks  FOR vbap-werks,
                  s_auart  FOR vbak-auart,
                  s_facdt  FOR vbap-zzfactoryreqdate,
                  s_erdat  FOR vbak-erdat NO-EXTENSION,
                  s_wadat  FOR likp-wadat_ist ,
                  s_ebeln  FOR ekko-ebeln,
                  s_ebelp  FOR ekpo-ebelp,
                  s_bsart  FOR ekko-bsart,
                  s_kvgr2  FOR knvv-kvgr2,
                  s_prodh  FOR vbap-prodh,
                  s_matkl  FOR mara-matkl,
                  s_mtart  FOR mara-mtart.

  PARAMETERS: p_yvalu TYPE i OBLIGATORY,
              p_xvalu TYPE i OBLIGATORY.
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-002.
  PARAMETERS : p_sales AS CHECKBOX,
               p_purc AS CHECKBOX.

SELECTION-SCREEN END OF BLOCK b2.

SELECTION-SCREEN BEGIN OF BLOCK b3 WITH FRAME TITLE TEXT-013.
  PARAMETERS : p_kvgr AS CHECKBOX.
SELECTION-SCREEN END OF BLOCK b3.
