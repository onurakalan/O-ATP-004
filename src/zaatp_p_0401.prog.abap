*&---------------------------------------------------------------------*
*& Report    : İlk Termine Uyum Performansı Raporu
*&---------------------------------------------------------------------*
*& Firma     : İmprova
*& Abap      : Onur Akalan
*& Modül     : Miraç İşbilir / Aybüke Aydemir
*& Tarih     : 19.01.2022
*&---------------------------------------------------------------------*
REPORT zaatp_p_0401.

INCLUDE :
 zaatp_i_0401_top,
 zaatp_i_0401_scr,
 zaatp_i_0001_excel,
 zaatp_i_0401_cls,
 zaatp_i_0401_mdl.


INITIALIZATION.
  gr_main = NEW #( ).

START-OF-SELECTION.
  gr_main->run( ).

END-OF-SELECTION.
  IF gr_main->mt_data IS INITIAL.
    MESSAGE s100(zaatp) DISPLAY LIKE 'E' RAISING no_data.
    EXIT.
  ENDIF.

  CALL SCREEN 100.
