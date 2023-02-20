from tkinter import *

from PIL import ImageTk, Image
from tkinter import messagebox, scrolledtext, ttk
import zpl, os, sys, win32print, serial 
import customtkinter, mysql.connector, time
import requests, json, threading, base64
from datetime import datetime, timedelta, date
from decimal import Decimal
import smtplib, ssl 
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


data = time.strftime('%d/%m/%Y', time.localtime())
hora = time.strftime('%H:%M:%S', time.localtime())
data_hoje = datetime.strptime(str(data), "%d/%m/%Y")

ano = ''
mes = ''
letra = ''

fonte_titulo = ("Calibri",13)
fonte_titulo_bold = ("Calibri Bold",13)
fonte_padrao = ("Calibri",12)
fonte_padrao_bold = ("Calibri Bold",12)
fonte_botoes = ("Calibri",12)
fonte_filtro = ("Calibri",9)
titulo = "Simec Etiquetas 1.4"

estilo_entry_padrao = {'justify':'center','font':fonte_padrao, 'fg_color':'#f3f3f3', 'bg_color':'#ffffff', 'text_color':'#2a2d2e', 'border_color':'#2a2d2e'}
estilo_entry_padrao_etiqueta = {'placeholder_text':" ", 'justify':'center','font':fonte_titulo_bold, 'fg_color':'#f3f3f3', 'bg_color':'#ffffff', 'text_color':'#1D366C', 'border_color':'#2a2d2e'}
estilo_optionmenu = {'fg_color':"#2c2c2c", 'button_color':'#1c1c1c', 'button_hover_color':'#4F4F4F', 'font':fonte_padrao, 'text_color':'#ffffff'}
estilo_botao_padrao_form = {'fg_color':'#232527', 'hover_color':'#343638', 'font':fonte_padrao, 'text_color':'#ffffff' }
estilo_botao_padrao_etiqueta = {'fg_color':'#1D366C', 'hover_color':'#1D366C', 'font':fonte_padrao, 'text_color':'#ffffff' }
estilo_botao_padrao_vermelho = {'fg_color':'#8B0000', 'hover_color':'#8B0000', 'font':fonte_padrao, 'text_color':'#ffffff' }
var_temp = ''
bitola = ''
aco = ''
tipo_produto = 0
usuario_logado = ''
status = ''
controle_loop = ''
desc1 = ''
desc2 = ''
desc3 = ''
status = 0

def setup_acesso():
    if usuario_logado[4] == 0: #OPERADOR
        btn_usuario.configure(state='disabled')
        btn_reimp.configure(state='normal')
        btn_avulsa.configure(state='normal')

    else:
        btn_usuario.configure(state='normal')
        btn_reimp.configure(state='normal')
        btn_avulsa.configure(state='normal')

def etiquetas():
    global controle_loop
    controle_loop = 'etiquetas'
    
    for widget in frame4.winfo_children():
        widget.destroy()

    def controle_lote():
        cursor.execute("SELECT * FROM prod_tipo WHERE nome=%s", (ent_prod_tipo.get(),))
        ult_seq = cursor.fetchone()[3]
        #print(tipo_produto, ult_seq)
        if ult_seq == None or ult_seq == '':
            ent_seq.configure(state='normal')
            ent_seq.delete(0, END)
            ent_seq.insert(0, '00001')
            ent_seq.configure(state='readonly')
            gerar_volume()
        else:
            lote = ult_seq[0:3]
            sequen = ult_seq[3:]
            #print(lote, sequen)
            nova_seq = int(sequen)+1
            ent_seq.configure(state='normal')
            ent_seq.delete(0,END)
            ent_seq.insert(0, str(nova_seq).zfill(5))
            ent_seq.configure(state='readonly')
            gerar_volume()
            if ent_lote.get()[0:3] != lote:
                ent_seq.configure(state='normal')
                ent_seq.delete(0, END)
                ent_seq.insert(0, '00001')
                ent_seq.configure(state='readonly')
                gerar_volume()

    def opt_op_clique(event):
        global tipo_produto
        cursor.execute("SELECT * FROM op WHERE numop=%s", (clique_op.get(),))
        prod_op = cursor.fetchone()
        global desc1
        desc1 = prod_op[8].strip()
        global desc2        
        desc2 = prod_op[9].strip()
        global desc3                
        desc3 = prod_op[10].strip()        

        if prod_op[4] == 'R':
            ent_prod_tipo.configure(state='normal')
            ent_prod_tipo.delete(0,END)
            ent_prod_tipo.insert(0, 'ROLO')
            ent_prod_tipo.configure(state='readonly')
            ent_produto.configure(state='normal')
            ent_produto.delete(0,END)
            ent_produto.insert(0, prod_op[8]+' '+ prod_op[9]+ ' ' + prod_op[10])
            ent_produto.configure(state='readonly')
            tipo_produto = 2

        elif prod_op[4] == 'V':
            ent_prod_tipo.configure(state='normal')
            ent_prod_tipo.delete(0,END)
            ent_prod_tipo.insert(0, 'BARRA')
            ent_prod_tipo.configure(state='readonly')
            ent_produto.configure(state='normal')
            ent_produto.delete(0,END)
            ent_produto.insert(0, prod_op[8]+' '+ prod_op[9]+ ' ' + prod_op[10])
            ent_produto.configure(state='readonly')
            tipo_produto = 1

        elif prod_op[4] == 'F':
            print(prod_op)
            ent_prod_tipo.configure(state='normal')
            ent_prod_tipo.delete(0,END)
            ent_prod_tipo.insert(0, 'FIO MAQUINA')
            ent_prod_tipo.configure(state='readonly')
            ent_produto.configure(state='normal')
            ent_produto.delete(0,END)
            #ent_produto.insert(0, f'{float(prod_op[2][7:]):.2f}'+' MM|ACO '+prod_op[2][2:6])
            ent_produto.insert(0, prod_op[8]+'|'+ prod_op[9])
            ent_produto.configure(state='readonly')
            tipo_produto = 3

        elif prod_op[4] == 'E':
            ent_prod_tipo.configure(state='normal')
            ent_prod_tipo.delete(0,END)
            ent_prod_tipo.insert(0, 'ENDIREITADO')
            ent_prod_tipo.configure(state='readonly')
            ent_produto.configure(state='normal')
            ent_produto.delete(0,END)
            ent_produto.insert(0, prod_op[8]+' '+ prod_op[9]+ ' ' + prod_op[10])
            ent_produto.configure(state='readonly')
            tipo_produto = 4

        ent_cod.configure(state='normal')
        ent_cod.delete(0,END)
        ent_cod.insert(0, prod_op[2])
        ent_cod.configure(state='readonly')
        controle_lote()
    
    def opt_corrida_clique(event):
        cursor.execute("SELECT * FROM corridas WHERE corrida=%s", (clique_corrida.get(),))
        corrida = cursor.fetchone()

        ent_qtd_pecas.configure(state='normal')
        ent_qtd_pecas.delete(0,END)
        ent_qtd_pecas.insert(0, corrida[2])
        ent_qtd_pecas.configure(state='readonly')
   
        ent_consumidas.configure(state='normal')
        ent_consumidas.delete(0,END)
        ent_consumidas.insert(0, corrida[4])
        ent_consumidas.configure(state='readonly')

        ent_total_perdas.configure(state='normal')
        ent_total_perdas.delete(0,END)
        ent_total_perdas.insert(0, corrida[3])
        ent_total_perdas.configure(state='readonly')

    def calculo_corridas(corrida, perdas):
        cursor.execute("select * from corridas where corrida=%s",(corrida,))
        resultado = cursor.fetchone()
        total_consumida = resultado[4]+1
        total_perda = resultado[3]
        if perdas != '':
            total_perda = resultado[3]+int(perdas)
        
        try:
            cursor.execute("UPDATE corridas SET\
                                qtd_perdas = %s,\
                                qtd_consumidas = %s\
                                WHERE corrida = %s", (total_perda, total_consumida, corrida,))
            db.commit()
        except Exception as e:
            messagebox.showerror('+Etiqueta:', 'Erro ao gravar as informações na tabela (corrida).', parent=root)
            print(e)
            return False
        
        cursor.execute("SELECT * FROM corridas WHERE corrida=%s", (corrida,))
        busca = cursor.fetchone()
        ent_qtd_pecas.configure(state='normal')
        ent_qtd_pecas.delete(0,END)
        ent_qtd_pecas.insert(0, busca[2])
        ent_qtd_pecas.configure(state='readonly')

        ent_consumidas.configure(state='normal')
        ent_consumidas.delete(0,END)
        ent_consumidas.insert(0, busca[4])
        ent_consumidas.configure(state='readonly')

        ent_total_perdas.configure(state='normal')
        ent_total_perdas.delete(0,END)
        ent_total_perdas.insert(0, busca[3])
        ent_total_perdas.configure(state='readonly')

    def setup_etiqueta():
        ent_data.insert(0, data)
        ent_data.configure(state='readonly')
        ent_lote.configure(state='readonly')
        ent_cod.configure(state='readonly')

    def gerar_volume():
        seq= ent_seq.get()
        letras = clique_op.get()[:3]
        print(letras+seq)
        '''cursor.execute("SELECT * FROM prod_tipo WHERE nome =%s", (ent_prod_tipo.get(),))
        letra = cursor.fetchone()[2]
        ano = ent_data.get()
        ano = ano[6:]
        ano = (chr(int(ano)-1950))
        mes = ent_data.get()
        mes = mes[3:-5]
        mes = (chr(int(mes)+64))'''
        ent_lote.configure(state='normal')
        ent_lote.delete(0,END)
        #ent_lote.insert(0,letra+ano+mes+seq)
        ent_lote.insert(0,letras+seq)        
        ent_lote.configure(state='readonly')

    def impressao():

        def confirma_impressao():
            cursor.execute("SELECT * FROM prod_tipo WHERE nome=%s", (ent_prod_tipo.get(),))
            produto = cursor.fetchone()[0]

            cursor.execute("SELECT * FROM producao WHERE lote= %s LIMIT 0,1", (lote,))
            verifica = cursor.fetchone()
            if verifica == None:
                try:
                    cursor.execute("INSERT INTO producao (\
                        data,\
                        hora,\
                        usuario,\
                        id_prod_tipo,\
                        prod_descricao,\
                        prod_codigo,\
                        corrida,\
                        lote,\
                        peso,\
                        envio_tentativa,\
                        op)\
                        values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", (data, hora, usuario, produto, prod_descricao, prod_codigo, corrida, lote, peso, 0, op))
                    db.commit()
                except Exception as e:
                    messagebox.showerror('+Etiqueta:', 'Erro ao gravar as informações na tabela (produção).', parent=root)
                    print(e)
                    return False
                try:
                    cursor.execute("UPDATE prod_tipo SET ultima_seq = %s WHERE id = %s", (lote,tipo_produto,))
                    db.commit()
                except Exception as e:
                    messagebox.showerror('+Etiqueta:', 'Erro ao registrar a última sequência (prod_tipo).', parent=root)
                    print(e)
                    return False


                def modelo1(): #Barra ou Endireitado
                    global desc1
                    global desc2
                    global desc3
                    l = zpl.Label(114,75)
                    pos_y = 30
                    pos_x = 2
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"{desc1}",
                        char_height=4,
                        char_width=4,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"{desc2.replace('MM','mm')}",
                        char_height=4,
                        char_width=4,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"{desc3.replace('M','m')}",
                        char_height=4,
                        char_width=4,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 6
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Codigo {prod_codigo}",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Volume {lote}",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()


                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Corrida {corrida}",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Peso {peso} kg",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 5
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"{norma_abnt}",
                        char_height=2,
                        char_width=2,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 3
                    pos_x +=10
                    l.origin(pos_x,pos_y+5)
                    l.write_text(
                        "Registro",
                        char_height=2,
                        char_width=2,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    l.origin(pos_x-2,pos_y+7)
                    l.write_text(
                        f"{inmetro}",
                        char_height=3,
                        char_width=2,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    l.origin(42, 17)
                    l.write_barcode(height=60, barcode_type='C', check_digit='Y', orientation='R')
                    l.write_text(f'{prod_codigo}{lote}{corrida}{peso}')
                    l.endorigin()

                    pos_y += 8
                    l.origin(12, pos_y+10)
                    l.write_barcode(height=80, barcode_type='C', check_digit='Y', orientation='N')
                    l.write_text(f'{lote}')
                    l.endorigin()

                    #print(height)

                    pg1 = l.dumpZPL()
                    #l.preview()

                    printer_name = win32print.GetDefaultPrinter ()
                    #
                    # raw_data could equally be raw PCL/PS read from
                    #  some print-to-file operation
                    #
                    if sys.version_info >= (3,):
                        #raw_data = bytes ("^XA^FO24,372^A0N,48,48^FB720,1,0,L,0^FDVERGALHAO CA-50^FS^FO24,420^A0N,48,48^FB720,1,0,L,0^FD8,00 MM ROLO^FS^FO24,468^A0N,48,48^FB720,1,0,L,0^FD//// CAUDA ////^FS^FO24,540^A0N,48,36^FB720,1,0,L,0^FDCodigo V1020B8.0^FS^FO24,588^A0N,48,36^FB720,1,0,L,0^FDVolume LHC14583^FS^FO24,636^A0N,48,36^FB720,1,0,L,0^FDCorrida 31244^FS^FO24,684^A0N,48,36^FB720,1,0,L,0^FDPeso 2,007 TON^FS^FO144,828^A0N,36,24^FB720,1,0,L,0^FDRegistro^FS^FO120,864^A0N,36,24^FB720,1,0,L,0^FD005713/2015^FS^FO504,204^BCB,60,N,N,Y,N^FDV1020B8.0 LHC1458331244 2,007^FS^FO144,996^BCN,80,S,N,Y,N^FDLHC14583^FS^XZ", "utf-8")
                        raw_data = bytes (pg1, "utf-8")
                        #print(raw_data) 
                    else:
                        raw_data = "This is a test"
                    hPrinter = win32print.OpenPrinter (printer_name)
                    try:
                        hJob = win32print.StartDocPrinter (hPrinter, 1, ("test of raw data", None, "RAW"))
                        try:
                            win32print.StartPagePrinter (hPrinter)
                            win32print.WritePrinter (hPrinter, raw_data)
                            win32print.EndPagePrinter (hPrinter)
                        finally:
                            win32print.EndDocPrinter (hPrinter)
                    finally:
                        win32print.ClosePrinter (hPrinter)

                def modelo2(): #Rolo
                    global desc1
                    global desc2
                    global desc3                    
                    l = zpl.Label(114,75)
                    pos_y = 31
                    pos_x = 2
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"{desc1}",
                        char_height=4,
                        char_width=4,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"{desc2.replace('MM','mm')}",
                        char_height=4,
                        char_width=4,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        "//// CABECA ////",
                        char_height=4,
                        char_width=4,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 6
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Codigo {prod_codigo}",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Volume {lote}",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()


                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Corrida {corrida}",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Peso {peso} kg",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 5
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"{norma_abnt}",
                        char_height=2,
                        char_width=2,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 2
                    pos_x +=10
                    l.origin(pos_x,pos_y+5)
                    l.write_text(
                        "Registro",
                        char_height=2,
                        char_width=2,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    l.origin(pos_x-2,pos_y+7)
                    l.write_text(
                        f"{inmetro}",
                        char_height=3,
                        char_width=2,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()


                    l.origin(42, 17)
                    l.write_barcode(height=60, barcode_type='C', check_digit='Y', orientation='R')
                    l.write_text(f'{prod_codigo}{lote}{corrida}{peso}')
                    l.endorigin()

                    pos_y += 8
                    l.origin(12, pos_y+10)
                    l.write_barcode(height=80, barcode_type='C', check_digit='Y', orientation='N')
                    l.write_text(f'{lote}')
                    l.endorigin()

                    #print(height)

                    #print(l.dumpZPL())
                    pg1 = l.dumpZPL()
                    pg2 = pg1.replace('//// CABECA ////', '//// CAUDA ////')
                    uniao_paginas = pg1+pg2
                    #print(uniao_paginas)
                    #l.preview()

                    printer_name = win32print.GetDefaultPrinter ()
                    #
                    # raw_data could equally be raw PCL/PS read from
                    #  some print-to-file operation
                    #
                    if sys.version_info >= (3,):
                        #raw_data = bytes ("^XA^FO24,372^A0N,48,48^FB720,1,0,L,0^FDVERGALHAO CA-50^FS^FO24,420^A0N,48,48^FB720,1,0,L,0^FD8,00 MM ROLO^FS^FO24,468^A0N,48,48^FB720,1,0,L,0^FD//// CAUDA ////^FS^FO24,540^A0N,48,36^FB720,1,0,L,0^FDCodigo V1020B8.0^FS^FO24,588^A0N,48,36^FB720,1,0,L,0^FDVolume LHC14583^FS^FO24,636^A0N,48,36^FB720,1,0,L,0^FDCorrida 31244^FS^FO24,684^A0N,48,36^FB720,1,0,L,0^FDPeso 2,007 TON^FS^FO144,828^A0N,36,24^FB720,1,0,L,0^FDRegistro^FS^FO120,864^A0N,36,24^FB720,1,0,L,0^FD005713/2015^FS^FO504,204^BCB,60,N,N,Y,N^FDV1020B8.0 LHC1458331244 2,007^FS^FO144,996^BCN,80,S,N,Y,N^FDLHC14583^FS^XZ", "utf-8")
                        raw_data = bytes (uniao_paginas, "utf-8")
                    else:
                        raw_data = "This is a test"
                    hPrinter = win32print.OpenPrinter (printer_name)
                    try:
                        hJob = win32print.StartDocPrinter (hPrinter, 1, ("test of raw data", None, "RAW"))
                        try:
                            win32print.StartPagePrinter (hPrinter)
                            win32print.WritePrinter (hPrinter, raw_data)
                            win32print.EndPagePrinter (hPrinter)
                        finally:
                            win32print.EndDocPrinter (hPrinter)
                    finally:
                        win32print.ClosePrinter (hPrinter)

                def modelo3(): #FM
                    global desc1
                    global desc2
                    global desc3

                    l = zpl.Label(114,75)
                    pos_y = 30
                    pos_x = 2
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"{desc1.replace('MM','mm')}",
                        char_height=4,
                        char_width=4,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()
                    #6.50 MM|AÇO 1060
                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"{desc2}",
                        char_height=4,
                        char_width=4,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        "//// CABECA ////",
                        char_height=4,
                        char_width=4,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 6
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Codigo {prod_codigo}",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Volume {lote}",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()


                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Corrida {corrida}",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Peso {peso} kg",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 5
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        "",
                        char_height=2,
                        char_width=2,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 3
                    pos_x +=10
                    l.origin(pos_x,pos_y+5)
                    l.write_text(
                        "",
                        char_height=2,
                        char_width=2,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    l.origin(pos_x-2,pos_y+7)
                    l.write_text(
                        f"{inmetro}",
                        char_height=3,
                        char_width=2,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    l.origin(42, 17)
                    l.write_barcode(height=60, barcode_type='C', check_digit='Y', orientation='R')
                    l.write_text(f'{prod_codigo}{lote}{corrida}{peso}')
                    l.endorigin()

                    pos_y += 8
                    l.origin(12, pos_y+10)
                    l.write_barcode(height=80, barcode_type='C', check_digit='Y', orientation='N')
                    l.write_text(f'{lote}')
                    l.endorigin()
                    
                    #l.preview()

                    #print(height)

                    pg1 = l.dumpZPL()
                    pg2 = pg1.replace('//// CABECA ////', '//// CAUDA ////')
                    uniao_paginas = pg1+pg2

                    printer_name = win32print.GetDefaultPrinter ()
                    #
                    # raw_data could equally be raw PCL/PS read from
                    #  some print-to-file operation
                    #
                    if sys.version_info >= (3,):
                        #raw_data = bytes ("^XA^FO24,372^A0N,48,48^FB720,1,0,L,0^FDVERGALHAO CA-50^FS^FO24,420^A0N,48,48^FB720,1,0,L,0^FD8,00 MM ROLO^FS^FO24,468^A0N,48,48^FB720,1,0,L,0^FD//// CAUDA ////^FS^FO24,540^A0N,48,36^FB720,1,0,L,0^FDCodigo V1020B8.0^FS^FO24,588^A0N,48,36^FB720,1,0,L,0^FDVolume LHC14583^FS^FO24,636^A0N,48,36^FB720,1,0,L,0^FDCorrida 31244^FS^FO24,684^A0N,48,36^FB720,1,0,L,0^FDPeso 2,007 TON^FS^FO144,828^A0N,36,24^FB720,1,0,L,0^FDRegistro^FS^FO120,864^A0N,36,24^FB720,1,0,L,0^FD005713/2015^FS^FO504,204^BCB,60,N,N,Y,N^FDV1020B8.0 LHC1458331244 2,007^FS^FO144,996^BCN,80,S,N,Y,N^FDLHC14583^FS^XZ", "utf-8")
                        raw_data = bytes (uniao_paginas, "utf-8")
                        #print(raw_data) 
                    else:
                        raw_data = "This is a test"
                    hPrinter = win32print.OpenPrinter (printer_name)
                    try:
                        hJob = win32print.StartDocPrinter (hPrinter, 1, ("test of raw data", None, "RAW"))
                        try:
                            win32print.StartPagePrinter (hPrinter)
                            win32print.WritePrinter (hPrinter, raw_data)
                            win32print.EndPagePrinter (hPrinter)
                        finally:
                            win32print.EndDocPrinter (hPrinter)
                    finally:
                        win32print.ClosePrinter (hPrinter)                    

                ent_data.configure(state='normal')
                ent_data.delete(0,END)
                ent_data.insert(0, data)
                ent_data.configure(state='readonly')                
                controle_lote()
                atualizar_lista_etiquetas()
                calculo_corridas(clique_corrida.get(), ent_perdas.get())
                ent_peso.focus_force()
                
                cursor.execute("SELECT inmetro from op WHERE numop = %s", (clique_op.get(),))
                etq = cursor.fetchone()

                cursor.execute("select norma_abnt from norma_abnt")
                norma_abnt = cursor.fetchone()[0]

                if etq == None or etq == '':
                    inmetro = ''
                else:
                    inmetro = etq[0]
                if tipo_produto == 1 or tipo_produto == 4:
                    imp = threading.Thread(target=modelo1)
                    imp.start()
                elif tipo_produto == 2:
                    imp = threading.Thread(target=modelo2)
                    imp.start()
                elif tipo_produto == 3:
                    imp = threading.Thread(target=modelo3)
                    imp.start()
                ent_perdas.delete(0,END)                
                ent_peso.delete(0,END)
            else:
                messagebox.showerror('+Etiqueta:', 'Lote já utilizado.', parent=root)

        data = ent_data.get()
        usuario = usuario_logado[2]
        prod_descricao = ent_produto.get()
        prod_codigo = ent_cod.get()
        corrida = clique_corrida.get()
        lote = ent_lote.get()
        peso = ent_peso.get()
        op = clique_op.get()
        int_perdas = ent_perdas.get().isnumeric()
        int_peso = peso.isnumeric()


        if clique_op.get() == '' or  corrida == '' or peso == '':
            messagebox.showwarning('+Etiqueta:', 'Todos os campos devem ser preenchidos.', parent=root)
        elif len(peso) > 4 or len(peso) <4 :
            messagebox.showwarning('+Etiqueta:', 'Peso inválido.\nNúmero de caracteres maior ou menor do que 4.', parent=root)
        elif ent_perdas.get() != '' and int_perdas == False:
            messagebox.showwarning('+Etiqueta:', 'Somente números inteiros são permitidos.', parent=root)        
        elif int_peso == False:
            messagebox.showwarning('+Etiqueta:', 'Somente números inteiros são permitidos.', parent=root)                
        else:
            cursor.execute("SELECT * FROM corridas WHERE corrida=%s", (clique_corrida.get(),))
            result_corrida = cursor.fetchone()
            restante = result_corrida[2]-result_corrida[3]
            if result_corrida[4] >= restante:
                r = messagebox.askyesno('+Etiqueta:', 'Limite de consumo atingido. Deseja continuar utilizando esta corrida?', parent=root)
                if r == True:
                    confirma_impressao()
                else:
                    return False
            else:
                confirma_impressao()

    def loop_etiquetas():
        if controle_loop == 'etiquetas': #// Loop para ficar atualizando a pagina etiquetas
            root.after(30000, atualizar_lista_etiquetas)

    def atualizar_lista_etiquetas():
        #print('lista_etiquetas')
        if controle_loop == 'etiquetas':
            db.cmd_reset_connection()
            tree_etiqueta.delete(*tree_etiqueta.get_children())
            cursor.execute("SELECT\
                A.id,\
                A.data,\
                A.hora,\
                A.usuario,\
                B.nome,\
                A.prod_descricao,\
                A.prod_codigo,\
                A.corrida,\
                A.lote,\
                A.peso,\
                IFNULL(A.envio_protheus,'') FROM producao A\
                inner join prod_tipo B on B.id = A.id_prod_tipo\
                where data = %s ORDER BY A.id DESC",(data,))
            cont = 0
            for row in cursor:
                if cont % 2 == 0:
                    tree_etiqueta.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1], row[10]),
                                            tags=('par',))
                else:
                    tree_etiqueta.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1], row[10]),
                                            tags=('impar',))
                cont += 1
                #print(row)
            loop_etiquetas()

    def thread_captura_peso():
        if bt_peso._state == NORMAL:
            a = threading.Thread(target=capturar_peso)
            a.start()
        else:
            pass

    def capturar_peso():
        def pega_peso():
            pausa_captura = 0
            contador = 1
            peso1 = '0'
            peso2 = '0'
            peso3 = '0' 
            while contador <= 3:
                ser = serial.Serial('COM1', 4800, timeout=1, bytesize=7)
                data = ser.read(50)
                #data = b'\x02*3    975   975\r\x02*3    975   975\r\x02*3    975   975'
                if data == b'':
                    messagebox.showwarning('+Capturar Peso:', 'Erro ao capturar o peso!\nVerifique a comunicação com a porta serial.\n\nConfig serial: COM1/4800/Bytesize=7', parent=root)
                    ent_peso.focus_force()
                    ser.close()
                    pausa_captura = 2
                    break                    
                else:
                    index_inicio_string = data.index(b'\r\x02*')
                    peso_capturado = data[index_inicio_string+3:index_inicio_string+17].decode()
                    peso_formatado = peso_capturado[4:9]
                    ser.close()
                    if contador == 1:
                        peso1 = peso_formatado
                    elif contador == 2:
                        peso2 = peso_formatado
                    elif contador == 3:
                        peso3 = peso_formatado
                    contador +=1

            lista=[peso1, peso2, peso3]
            #lista=['2540', '2541', '2541']
            lista_nova = []
            if peso1 != '0' and peso2 != '0' and peso3 != '0':
                for x in lista:
                    repetido = 0    
                    for y in lista:
                        if x == y:
                            repetido +=1
                            if repetido >= 2:
                                lista_nova.append(x)
            if len(lista_nova) != 0:
                return lista_nova[0].replace(" ", "")                
            else:
                if pausa_captura == 2:
                    return 2
                else:
                    return 0

        op = clique_op.get()
        corrida = clique_corrida.get()

        ent_peso.delete(0,END)
        ent_peso.insert(0,'Capturando o peso! Aguarde.')
        bt_peso.configure(state=DISABLED)

        if op == '' or corrida == '':
            messagebox.showwarning('+Capturar Peso:', 'Selecione a Ordem de Produção ou Corrida.', parent=root)
        else:
            if pega_peso() == 2:
                ent_peso.delete(0,END)
                ent_peso.focus_force()
                bt_peso.configure(state=NORMAL)
            elif pega_peso() == 0:
                tentativa_pega_peso = 0
                while tentativa_pega_peso <1:
                    if pega_peso() == 0:
                        tentativa_pega_peso += 1                    
                    else:
                        ent_peso.delete(0,END)
                        ent_peso.insert(0, pega_peso())
                        bt_peso.configure(state=NORMAL)
                        break
                if tentativa_pega_peso == 1:
                    ent_peso.delete(0,END)
                    ent_peso.insert(0, pega_peso())
                    ent_peso.focus_force()
                    bt_peso.configure(state=NORMAL)
            else:
                ent_peso.delete(0,END)
                ent_peso.insert(0, pega_peso())
                bt_peso.configure(state=NORMAL)

    def encerrar_corrida():
        corrida = clique_corrida.get()
        lista_corrida_nova = []
        if corrida != '':
            r = messagebox.askyesno('+Encerrar Corrida:', f'Confirma o encerramento da corrida {corrida}?', parent=root)
            if r == True:
                opt_corrida.set('')                
                cursor.execute("UPDATE corridas SET encerrada = 'Sim' WHERE corrida = %s",(corrida,))
                db.commit()
                time.sleep(1)
                cursor.execute("SELECT * FROM corridas where IFNULL(encerrada, '') <> 'Sim' ORDER BY ID DESC")
                for i in cursor:
                    if (i[2] - i[3]) > i[4]:
                        lista_corrida_nova.append(i[1])
                opt_corrida.configure(values=lista_corrida_nova)
                
    def encerrar_corrida_thread():
        s = threading.Thread(target=encerrar_corrida)
        s.start()
    fr1 = Frame(frame4, bg='#ffffff')
    fr1.pack(side=TOP, fill=X)
    fr2 = Frame(frame4, bg='#222222')  # linha
    fr2.pack(side=TOP, fill=X, pady=5)
    fr3 = Frame(frame4, bg='#ffffff')
    fr3.pack(side=TOP, fill=X)
    fr4 = Frame(frame4, bg='#222222')  # linha
    fr4.pack(side=TOP, fill=X, pady=5)
    fr5 = Frame(frame4, bg='#ffffff')
    fr5.pack(side=TOP, fill=X)
    fr5_0 = Frame(frame4, bg='#222222')  # linha
    fr5_0.pack(side=TOP, fill=X, pady=5)    
    fr5_1 = Frame(frame4, bg='#ffffff')
    fr5_1.pack(side=TOP, fill=X, pady=5)    
    fr5_2 = Frame(frame4, bg='#ffffff')
    fr5_2.pack(side=TOP, fill=X, pady=5)    
    fr6 = Frame(frame4, bg='#ffffff')
    fr6.pack(side=TOP, fill=X)    
    fr7 = Frame(frame4, bg='#222222')  # linha
    fr7.pack(side=TOP, fill=X, pady=5)
    fr8 = Frame(frame4, bg='#ffffff')
    fr8.pack(side=TOP, fill=X)    
    fr9 = Frame(frame4, bg='#222222')  # linha
    fr9.pack(side=TOP, fill=X, pady=5)
    fr10 = Frame(frame4, bg='#ffffff')
    fr10.pack(side=TOP, fill=BOTH, expand=TRUE)    


    #//////FRAME1
    image_invent = Image.open('img\\impressora.png')
    resize_invent = image_invent.resize((60, 40))
    nova_image_invent = ImageTk.PhotoImage(resize_invent)
    lbl = Label(fr1, text=" Etiquetas:", image=nova_image_invent, compound="left", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
    lbl.grid(row=0, column=1)
    fr1.grid_columnconfigure(0, weight=1)
    fr1.grid_columnconfigure(2, weight=1)
    #//////FRAME2 LINHA
    
    #//////FRAME3
    fr3.grid_columnconfigure(0, weight=1)
    fr3.grid_columnconfigure(6, weight=1)

    lbl = Label(fr3, text="Data:", font=fonte_padrao, fg='#222222', bg='#ffffff')
    lbl.grid(row=0, column=1, sticky=W, padx=5)
    ent_data = customtkinter.CTkEntry(fr3, **estilo_entry_padrao, width=140)
    ent_data.grid(row=1, column=1, padx=5)

    clique_op = StringVar()
    lista_op = []
    cursor.execute("SELECT * FROM op LIMIT 0,1")
    result = cursor.fetchone()
    if result == None:
        lista_op.append('')
    else:
        cursor.execute("SELECT * FROM op where encerrado IS NULL ORDER BY numop")
        for i in cursor:
            ano_protheus_fim =(i[6])[:4]
            mes_protheus_fim =(i[6])[4:6]
            dia_protheus_fim =(i[6])[6:8]
            data_protheus_fim = dia_protheus_fim+'/'+mes_protheus_fim+'/'+ano_protheus_fim
            data_op_fim = datetime.strptime(str(data_protheus_fim), "%d/%m/%Y")

            ano_protheus_ini =(i[5])[:4]
            mes_protheus_ini =(i[5])[4:6]
            dia_protheus_ini =(i[5])[6:8]
            data_protheus_ini = dia_protheus_ini+'/'+mes_protheus_ini+'/'+ano_protheus_ini
            data_op_ini = datetime.strptime(str(data_protheus_ini), "%d/%m/%Y")

            if data_hoje >= data_op_ini and data_hoje <=data_op_fim:
                lista_op.append(i[1])


    lbl=Label(fr3, text='Ordem de Produção: ', font=fonte_padrao_bold, bg='#ffffff', fg='#1D366C')
    lbl.grid(row=0, column=2, sticky="w", padx=5)
    #opt_op = customtkinter.CTkOptionMenu(fr3, variable=clique_op, values=lista_op, width=200, **estilo_optionmenu,command=opt_op_clique)
    #opt_op.grid(row=1, column=2, padx=5)

    opt_op = ttk.Combobox(fr3, textvariable=clique_op, values=lista_op, width=22, state='readonly', font=fonte_padrao, foreground='#2a2d2e', justify='center',)
    opt_op.grid(row=1, column=2, padx=5)
    opt_op.bind("<<ComboboxSelected>>", opt_op_clique)


    clique_corrida = StringVar()
    lista_corrida = []
    cursor.execute("SELECT * FROM corridas LIMIT 0,1")
    result = cursor.fetchone()
    if result == None:
        lista_corrida.append('')
    else:
        cursor.execute("SELECT * FROM corridas where IFNULL(encerrada, '') <> 'Sim' ORDER BY ID DESC")
        for i in cursor:
            if (i[2] - i[3]) > i[4]:
                lista_corrida.append(i[1])
    lbl=Label(fr3, text='Corrida: ', font=fonte_padrao_bold, bg='#ffffff', fg='#1D366C')
    lbl.grid(row=0, column=3, sticky="w", padx=5)
    #opt_corrida = customtkinter.CTkOptionMenu(fr3, variable=clique_corrida, values=lista_corrida, width=200, **estilo_optionmenu, command=opt_corrida_clique)
    #opt_corrida.grid(row=1, column=3, padx=5)

    opt_corrida = ttk.Combobox(fr3, textvariable=clique_corrida, values=lista_corrida, width=22, state='readonly', font=fonte_padrao, foreground='#2a2d2e', justify='center',)
    opt_corrida.grid(row=1, column=3, padx=5)
    opt_corrida.bind("<<ComboboxSelected>>", opt_corrida_clique)
    
    #//////FRAME4 LINHA

    #//////FRAME5
    lbl=Label(fr5, text='Informações sobre a corrida: ', font=fonte_padrao_bold, bg='#ffffff', fg='#1D366C')
    lbl.grid(row=0, column=1, sticky="w", padx=10, columnspan=4)

    lbl=Label(fr5, text='Qtd. de Peças: ', font=fonte_padrao, bg='#ffffff', fg='#000000')
    lbl.grid(row=1, column=1, sticky="w", padx=10)
    ent_qtd_pecas = customtkinter.CTkEntry(fr5, width=130, **estilo_entry_padrao)
    ent_qtd_pecas.grid(row=2, column=1, sticky="w", padx=10)
    ent_qtd_pecas.configure(state='readonly')

    lbl=Label(fr5, text='Qtd. Consumida: ', font=fonte_padrao, bg='#ffffff', fg='#000000')
    lbl.grid(row=1, column=2, sticky="w", padx=5)
    ent_consumidas = customtkinter.CTkEntry(fr5, width=130, **estilo_entry_padrao)
    ent_consumidas.grid(row=2, column=2, sticky="w", padx=5)
    ent_consumidas.configure(state='readonly')

    lbl=Label(fr5, text='Total de Perdas: ', font=fonte_padrao, bg='#ffffff', fg='#000000')
    lbl.grid(row=1, column=3, sticky="w", padx=5)
    ent_total_perdas = customtkinter.CTkEntry(fr5, width=130, **estilo_entry_padrao)
    ent_total_perdas.grid(row=2, column=3, sticky="w", padx=5)
    ent_total_perdas.configure(state='readonly')

    lbl=Label(fr5, text='Informe as Perdas: ', font=fonte_padrao_bold, bg='#ffffff', fg='#880000')
    lbl.grid(row=1, column=4, sticky="w", padx=10)
    ent_perdas = customtkinter.CTkEntry(fr5, width=130, **estilo_entry_padrao)
    ent_perdas.grid(row=2, column=4, sticky="w", padx=10)

    fr5.grid_columnconfigure(0, weight=1)
    fr5.grid_columnconfigure(5, weight=1)

    #//////FRAME5_0 LINHA

    #//////FRAME5_1
    lbl=Label(fr5_1, text='Linha de Produção: ', font=fonte_padrao, bg='#ffffff', fg='#000000')
    lbl.grid(row=0, column=1, sticky="w", padx=10)
    ent_prod_tipo = customtkinter.CTkEntry(fr5_1, width=210, **estilo_entry_padrao)
    ent_prod_tipo.grid(row=1, column=1, padx=10)
    ent_prod_tipo.configure(state='readonly')

    lbl=Label(fr5_1, text='Produto:', font=fonte_padrao, bg='#ffffff', fg='#000000')
    lbl.grid(row=0, column=2, sticky="w")
    ent_produto = customtkinter.CTkEntry(fr5_1, width=345, **estilo_entry_padrao)
    ent_produto.grid(row=1, column=2)
    ent_produto.configure(state='readonly')

    fr5_1.grid_columnconfigure(0, weight=1)
    fr5_1.grid_columnconfigure(3, weight=1)

    #//////FRAME5_2
    lbl = Label(fr5_2, text="Código:", font=fonte_padrao, fg='#000000', bg='#ffffff')
    lbl.grid(row=2, column=1, sticky=W, padx=5)
    ent_cod = customtkinter.CTkEntry(fr5_2, **estilo_entry_padrao, width=180)
    ent_cod.grid(row=3, column=1, sticky=W, padx=5)

    lbl = Label(fr5_2, text="Lote|Volume:", font=fonte_padrao, fg='#000000', bg='#ffffff')
    lbl.grid(row=2, column=2, sticky=W, padx=5)
    ent_lote = customtkinter.CTkEntry(fr5_2, **estilo_entry_padrao, width=180)
    ent_lote.grid(row=3, column=2, sticky=W, padx=5)

    lbl = Label(fr5_2, text="Sequência:", font=fonte_padrao, fg='#000000', bg='#ffffff')
    lbl.grid(row=2, column=3, sticky=W, padx=5)
    ent_seq = customtkinter.CTkEntry(fr5_2, **estilo_entry_padrao, width=180)
    ent_seq.grid(row=3, column=3, sticky=W, padx=5)
    ent_seq.configure(state='readonly')

    fr5_2.grid_columnconfigure(0, weight=1)
    fr5_2.grid_columnconfigure(4, weight=1)

    #//////FRAME6
    fr6.grid_columnconfigure(0, weight=1)
    fr6.grid_columnconfigure(6, weight=1)

    lbl = Label(fr6, text="Peso (KG):", font=fonte_padrao_bold, fg='#1D366C', bg='#ffffff')
    lbl.grid(row=0, column=1, sticky=W, padx=5)
    ent_peso = customtkinter.CTkEntry(fr6, **estilo_entry_padrao, width=200)
    ent_peso.grid(row=1, column=1, sticky=W, padx=5)

    bt_peso = customtkinter.CTkButton(fr6, text='Capturar|Peso (F2)', **estilo_botao_padrao_form, command=thread_captura_peso)
    bt_peso.grid(row=1, column=2)
    

    #//////FRAME7 LINHA
    
    #//////FRAME8
    fr8.grid_columnconfigure(0, weight=1)
    fr8.grid_columnconfigure(6, weight=1)
    global bt_imp
    bt_imp = customtkinter.CTkButton(fr8, text='Imprimir (F1)', **estilo_botao_padrao_etiqueta, command=impressao)
    bt_imp.grid(row=1, column=1, padx=30)
    bt_imp = customtkinter.CTkButton(fr8, text='Encerrar Corrida?', **estilo_botao_padrao_vermelho, command=encerrar_corrida_thread)
    bt_imp.grid(row=1, column=2)    
    root.bind('<F1>', lambda event: impressao())
    root.bind('<F2>', lambda event: thread_captura_peso())
    #//////FRAME9 LINHA
    
    #//////FRAME10
    style = ttk.Style()
    # style.theme_use('default')
    style.configure('Treeview',
                    background='#ffffff',
                    rowheight=24,
                    fieldbackground='#ffffff',
                    font=fonte_padrao)
    style.configure("Treeview.Heading",
                    foreground='#1d366c',
                    background="#ffffff",
                    font=fonte_padrao_bold)
    style.map('Treeview', background=[('selected', '#1D366C')])

    tree_etiqueta = ttk.Treeview(fr10, selectmode='browse')
    vsb = ttk.Scrollbar(fr10, orient="vertical", command=tree_etiqueta.yview)
    vsb.pack(side=RIGHT, fill='y')
    tree_etiqueta.configure(yscrollcommand=vsb.set)
    vsbx = ttk.Scrollbar(fr10, orient="horizontal", command=tree_etiqueta.xview)
    vsbx.pack(side=BOTTOM, fill='x')
    tree_etiqueta.configure(xscrollcommand=vsbx.set)
    tree_etiqueta.pack(side=LEFT, fill=BOTH, expand=True, anchor='n')
    tree_etiqueta["columns"] = ("1", "2", "3", "4", "5", "6", "7", "8")
    tree_etiqueta['show'] = 'headings'
    tree_etiqueta.column("#1", anchor='c', minwidth=50, width=100, stretch = True)
    tree_etiqueta.column("#2", anchor='c', minwidth=50, width=100, stretch = True)
    tree_etiqueta.column("#3", anchor='c', minwidth=50, width=100, stretch = True)
    tree_etiqueta.column("#4", anchor='c', minwidth=50, width=100, stretch = True)
    tree_etiqueta.column("#5", anchor='c', minwidth=50, width=100, stretch = True)
    tree_etiqueta.column("#6", anchor='c', minwidth=50, width=100, stretch = True)
    tree_etiqueta.column("#7", anchor='c', minwidth=50, width=100, stretch = True)
    tree_etiqueta.column("#8", anchor='c', minwidth=50, width=100, stretch = True)
    tree_etiqueta.heading("1", text="Registro DB")
    tree_etiqueta.heading("2", text="Produto")
    tree_etiqueta.heading("3", text="Usuário")
    tree_etiqueta.heading("4", text="Volume")
    tree_etiqueta.heading("5", text="Corrida")
    tree_etiqueta.heading("6", text="Peso")
    tree_etiqueta.heading("7", text="Data")
    tree_etiqueta.heading("8", text="Envio Protheus")    
    tree_etiqueta.tag_configure('par', background='#e9e9e9')
    tree_etiqueta.tag_configure('impar', background='#ffffff')
    #tree_etiqueta.bind("<Double-1>", duploclique_tree_etiqueta)
    fr10.grid_columnconfigure(0, weight=1)
    fr10.grid_columnconfigure(3, weight=1)

    atualizar_lista_etiquetas()
    setup_etiqueta()
    mainloop()

def avulsas():
    global controle_loop
    controle_loop = ''
        
    for widget in frame4.winfo_children():
        widget.destroy()
    
    def barra_endireitado():
        for widget in fr3.winfo_children():
            widget.destroy()
        for widget in fr5.winfo_children():
            widget.destroy()

        def cor_out_lbl_1(event):
            lbl_1.configure(fg='#222222')
            verifica = ent_campo1.get().replace('.', ',')
            cursor.execute("select * from inmetro where bitola = %s",(verifica,))
            result = cursor.fetchone()
            if result != None:
                lbl_8.configure(text=f"Registro\n{result[2]}")
                ent_campo8.configure(state='normal')
                ent_campo8.delete(0, END)
                ent_campo8.insert(0, result[2])
                ent_campo8.configure(state='readonly')
            else:
                messagebox.showerror('+Etiqueta:', 'Atenção! Bitola incorreta.\nDigite o valor e o padrão correto (0,00)', parent=root)
                ent_campo1.delete(0,END)
                ent_campo8.configure(state='normal')
                ent_campo8.delete(0, END)
                ent_campo8.configure(state='readonly')


        def cor_in_lbl_1(event):
            lbl_1.configure(fg='#0275d8')
        def key_press_lbl_1(event):
            captura = ent_campo1.get()
            lbl_1.configure(text=f'{captura.replace(".", ",")} mm BARRA')

        def cor_out_lbl_2(event):
            lbl_2.configure(fg='#222222')
        def cor_in_lbl_2(event):
            lbl_2.configure(fg='#0275d8')
        def key_press_lbl_2(event):
            captura = ent_campo2.get()
            lbl_2.configure(text=f'{captura.replace(".", ",")} m (S)')

        def cor_out_lbl_3(event):
            lbl_3.configure(fg='#222222')
        def cor_in_lbl_3(event):
            lbl_3.configure(fg='#0275d8')
        def key_press_lbl_3(event):
            captura = ent_campo3.get()
            lbl_3.configure(text=f'Codigo {captura.replace(",", ".").upper()}')

        def cor_out_lbl_4(event):
            lbl_4.configure(fg='#222222')
        def cor_in_lbl_4(event):
            lbl_4.configure(fg='#0275d8')
        def key_press_lbl_4(event):
            captura = ent_campo4.get()
            lbl_4.configure(text=f'Volume {captura.replace(",", ".").upper()}')

        def cor_out_lbl_5(event):
            lbl_5.configure(fg='#222222')
        def cor_in_lbl_5(event):
            lbl_5.configure(fg='#0275d8')
        def key_press_lbl_5(event):
            captura = ent_campo5.get()
            lbl_5.configure(text=f'Corrida {captura}')

        def cor_out_lbl_6(event):
            lbl_6.configure(fg='#222222')
        def cor_in_lbl_6(event):
            lbl_6.configure(fg='#0275d8')
        def key_press_lbl_6(event):
            captura = ent_campo6.get()
            lbl_6.configure(text=f'Peso {captura} kg')

        def cor_out_lbl_8(event):
            lbl_8.configure(fg='#222222')
        def cor_in_lbl_8(event):
            lbl_8.configure(fg='#0275d8')
        def key_press_lbl_8(event):
            captura = ent_campo8.get()
            lbl_8.configure(text=f'Registro\n{captura}', justify='center')

        def impressao():
            def modelo1(): #Barra ou Endireitado
                l = zpl.Label(114,75)
                pos_y = 30
                pos_x = 2
                l.origin(pos_x,pos_y)
                l.write_text(
                    "VERGALHAO CA-50",
                    char_height=4,
                    char_width=4,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 4
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"{desc1} mm BARRA",
                    char_height=4,
                    char_width=4,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 4
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"{desc2} m (S)",
                    char_height=4,
                    char_width=4,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 6
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"Codigo {codigo}",
                    char_height=4,
                    char_width=3,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 4
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"Volume {volume}",
                    char_height=4,
                    char_width=3,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()


                pos_y += 4
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"Corrida {corrida}",
                    char_height=4,
                    char_width=3,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 4
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"Peso {peso} kg",
                    char_height=4,
                    char_width=3,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 5
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"{abnt}",
                    char_height=2,
                    char_width=2,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 3
                pos_x +=10
                l.origin(pos_x,pos_y+5)
                l.write_text(
                    "Registro",
                    char_height=2,
                    char_width=2,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                l.origin(pos_x-2,pos_y+7)
                l.write_text(
                    f"{inmetro}",
                    char_height=3,
                    char_width=2,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                l.origin(42, 17)
                l.write_barcode(height=60, barcode_type='C', check_digit='Y', orientation='R')
                l.write_text(f'{codigo}{volume}{corrida}{peso}')
                l.endorigin()

                pos_y += 8
                l.origin(12, pos_y+10)
                l.write_barcode(height=80, barcode_type='C', check_digit='Y', orientation='N',)
                l.write_text(f'{volume}')
                l.endorigin()

                #print(height)

                pg1 = l.dumpZPL()
                #l.preview()

                printer_name = win32print.GetDefaultPrinter ()
                #
                # raw_data could equally be raw PCL/PS read from
                #  some print-to-file operation
                #
                if sys.version_info >= (3,):
                    #raw_data = bytes ("^XA^FO24,372^A0N,48,48^FB720,1,0,L,0^FDVERGALHAO CA-50^FS^FO24,420^A0N,48,48^FB720,1,0,L,0^FD8,00 MM ROLO^FS^FO24,468^A0N,48,48^FB720,1,0,L,0^FD//// CAUDA ////^FS^FO24,540^A0N,48,36^FB720,1,0,L,0^FDCodigo V1020B8.0^FS^FO24,588^A0N,48,36^FB720,1,0,L,0^FDVolume LHC14583^FS^FO24,636^A0N,48,36^FB720,1,0,L,0^FDCorrida 31244^FS^FO24,684^A0N,48,36^FB720,1,0,L,0^FDPeso 2,007 TON^FS^FO144,828^A0N,36,24^FB720,1,0,L,0^FDRegistro^FS^FO120,864^A0N,36,24^FB720,1,0,L,0^FD005713/2015^FS^FO504,204^BCB,60,N,N,Y,N^FDV1020B8.0 LHC1458331244 2,007^FS^FO144,996^BCN,80,S,N,Y,N^FDLHC14583^FS^XZ", "utf-8")
                    raw_data = bytes (pg1, "utf-8")
                    #print(raw_data) 
                else:
                    raw_data = "This is a test"
                hPrinter = win32print.OpenPrinter (printer_name)
                try:
                    hJob = win32print.StartDocPrinter (hPrinter, 1, ("test of raw data", None, "RAW"))
                    try:
                        win32print.StartPagePrinter (hPrinter)
                        win32print.WritePrinter (hPrinter, raw_data)
                        win32print.EndPagePrinter (hPrinter)
                    finally:
                        win32print.EndDocPrinter (hPrinter)
                finally:
                    win32print.ClosePrinter (hPrinter)

            def limpa_campos():
                ent_campo1.delete(0,END)
                ent_campo2.delete(0,END)
                ent_campo3.delete(0,END)
                ent_campo4.delete(0,END)
                ent_campo5.delete(0,END)
                ent_campo6.delete(0,END)
                ent_campo8.configure(state='normal')
                ent_campo8.delete(0,END)
                ent_campo8.configure(state='readonly')

            tipo_produto = 'BARRA/ENDIREITADO'
            desc1 = ent_campo1.get().replace('.', ',')
            desc2 = ent_campo2.get().replace('.', ',')
            codigo = ent_campo3.get().upper().replace(',', '.')
            volume = ent_campo4.get().upper()
            corrida = ent_campo5.get()
            peso = ent_campo6.get()
            inmetro = ent_campo8.get()
            usuario = usuario_logado[2]
            if desc1=='' or desc2 =='' or codigo =='' or volume ==''or corrida ==''or peso ==''or inmetro =='':
                messagebox.showwarning('+Etiqueta:', 'Todos os campos devem ser preenchidos.', parent=root)
            else:
                try:
                    cursor.execute("INSERT INTO etiquetas_avulsas (\
                        data,\
                        hora,\
                        usuario,\
                        tipo_produto,\
                        desc1,\
                        desc2,\
                        codigo,\
                        volume,\
                        corrida,\
                        peso,\
                        abnt,\
                        inmetro)\
                        values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", (data, hora, usuario, tipo_produto, desc1, desc2, codigo, volume, corrida, peso, abnt, inmetro))
                    db.commit()
                except Exception as e:
                    messagebox.showerror('+Etiqueta:', 'Erro ao gravar as informações na tabela (etiquetas_avulsas).', parent=root)
                    print(e)
                    return False
                imp = threading.Thread(target=modelo1)
                imp.start()
                limpa_campos()
                ent_campo2.focus_force()

        fr3_1 = Frame(fr3, bg='#ffffff')
        fr3_1.pack(side=LEFT, fill=X, expand=True)
        fr3_2 = Frame(fr3, bg='#ffffff')
        fr3_2.pack(side=RIGHT, fill=X, expand=True)            
        
        #//////FRAME3 e 3_1
        cursor.execute("SELECT norma_abnt FROM simec_etiquetas.norma_abnt")
        abnt = cursor.fetchone()[0]
        image_avulsas = Image.open('img\\avulsas.png')
        resize_avulsas = image_avulsas.resize((290, 114))
        nova_image_avulsas = ImageTk.PhotoImage(resize_avulsas)
        lbl = Label(fr3_1, image=nova_image_avulsas, bg='#ffffff')
        lbl.grid(row=0, column=1)
        lbl = Label(fr3_1, text="VERGALHAO CA-50", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=1, column=1, sticky='W', padx=10)
        lbl_1 = Label(fr3_1, text="0,00 mm BARRA", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_1.grid(row=2, column=1, sticky='W', padx=10)    
        lbl_2 = Label(fr3_1, text="0,00 m (S)", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_2.grid(row=3, column=1, sticky='W', padx=10)    
        lbl_3 = Label(fr3_1, text="Codigo VCA00000", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_3.grid(row=4, column=1, sticky='W', padx=10, pady=(6,0))    
        lbl_4 = Label(fr3_1, text="Volume LIA00000", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_4.grid(row=5, column=1, sticky='W', padx=10)    
        lbl_5 = Label(fr3_1, text="Corrida 00000", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_5.grid(row=6, column=1, sticky='W', padx=10)    
        lbl_6 = Label(fr3_1, text="Peso 0000 kg", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_6.grid(row=7, column=1, sticky='W', padx=10)    
        lbl_7 = Label(fr3_1, text=abnt, font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl_7.grid(row=8, column=1, sticky='W', padx=10)    
        lbl_8 = Label(fr3_1, text="Registro\n000000/0000", font=fonte_padrao_bold, fg='#222222', bg='#ffffff', justify='center')
        lbl_8.grid(row=9, column=1, sticky='W', padx=(50,0), pady=(5,0))    
        
        fr3_1.grid_columnconfigure(0, weight=1)
        
        #//////FRAME3 e 3_2
        lbl = Label(fr3_2, text="Desc.1:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=1, column=2)
        ent_campo1 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo1.grid(row=1, column=3, pady=(10,8))
        ent_campo1.bind('<KeyRelease>',key_press_lbl_1)
        ent_campo1.bind("<FocusIn>", cor_in_lbl_1)
        ent_campo1.bind("<FocusOut>", cor_out_lbl_1)

        lbl = Label(fr3_2, text="Desc.2:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=2, column=2)
        ent_campo2 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo2.grid(row=2, column=3, pady=(8))
        ent_campo2.bind('<KeyRelease>',key_press_lbl_2)
        ent_campo2.bind("<FocusIn>", cor_in_lbl_2)
        ent_campo2.bind("<FocusOut>", cor_out_lbl_2)

        lbl = Label(fr3_2, text="Codigo:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=3, column=2)
        ent_campo3 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo3.grid(row=3, column=3, pady=(8))
        ent_campo3.bind('<KeyRelease>',key_press_lbl_3)
        ent_campo3.bind("<FocusIn>", cor_in_lbl_3)
        ent_campo3.bind("<FocusOut>", cor_out_lbl_3)

        lbl = Label(fr3_2, text="Volume:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=4, column=2)
        ent_campo4 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo4.grid(row=4, column=3, pady=(8))
        ent_campo4.bind('<KeyRelease>',key_press_lbl_4)
        ent_campo4.bind("<FocusIn>", cor_in_lbl_4)
        ent_campo4.bind("<FocusOut>", cor_out_lbl_4)

        lbl = Label(fr3_2, text="Corrida:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=5, column=2)
        ent_campo5 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo5.grid(row=5, column=3, pady=(8))
        ent_campo5.bind('<KeyRelease>',key_press_lbl_5)
        ent_campo5.bind("<FocusIn>", cor_in_lbl_5)
        ent_campo5.bind("<FocusOut>", cor_out_lbl_5)

        lbl = Label(fr3_2, text="Peso:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=6, column=2)
        ent_campo6 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo6.grid(row=6, column=3, pady=(8))
        ent_campo6.bind('<KeyRelease>',key_press_lbl_6)
        ent_campo6.bind("<FocusIn>", cor_in_lbl_6)
        ent_campo6.bind("<FocusOut>", cor_out_lbl_6)

        lbl = Label(fr3_2, text="Inmetro:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=7, column=2)
        ent_campo8 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo8.grid(row=7, column=3, pady=(8))
        ent_campo8.bind('<KeyRelease>',key_press_lbl_8)
        ent_campo8.bind("<FocusIn>", cor_in_lbl_8)
        ent_campo8.bind("<FocusOut>", cor_out_lbl_8)
        ent_campo8.configure(state='readonly')
        
        fr3_2.grid_columnconfigure(4, weight=1)
        
        #//////FRAME5
        fr5.grid_columnconfigure(0, weight=1)
        fr5.grid_columnconfigure(6, weight=1)
        global bt_imp
        bt_imp = customtkinter.CTkButton(fr5, text='Imprimir (F1)', **estilo_botao_padrao_etiqueta, command=impressao)
        bt_imp.grid(row=1, column=1)
        root.bind('<F1>', lambda event: impressao())        
        mainloop()    

    def rolo():
        for widget in fr3.winfo_children():
            widget.destroy()
        for widget in fr5.winfo_children():
            widget.destroy()

        def cor_out_lbl_1(event):
            lbl_1.configure(fg='#222222')
            verifica = ent_campo1.get().replace('.', ',')
            cursor.execute("select * from inmetro where bitola = %s",(verifica,))
            result = cursor.fetchone()
            if result != None:
                lbl_8.configure(text=f"Registro\n{result[2]}")
                ent_campo8.configure(state='normal')
                ent_campo8.delete(0, END)
                ent_campo8.insert(0, result[2])
                ent_campo8.configure(state='readonly')
            else:
                messagebox.showerror('+Etiqueta:', 'Atenção! Bitola incorreta.\nDigite o valor e o padrão correto (0,00)', parent=root)
                ent_campo1.delete(0,END)
                ent_campo8.configure(state='normal')
                ent_campo8.delete(0, END)
                ent_campo8.configure(state='readonly')

        def cor_in_lbl_1(event):
            lbl_1.configure(fg='#0275d8')
        def key_press_lbl_1(event):
            captura = ent_campo1.get()
            lbl_1.configure(text=f'{captura.replace(".", ",")} mm ROLO (S)')

        def cor_out_lbl_3(event):
            lbl_3.configure(fg='#222222')
        def cor_in_lbl_3(event):
            lbl_3.configure(fg='#0275d8')
        def key_press_lbl_3(event):
            captura = ent_campo3.get()
            lbl_3.configure(text=f'Codigo {captura.replace(",", ".").upper()}')

        def cor_out_lbl_4(event):
            lbl_4.configure(fg='#222222')
        def cor_in_lbl_4(event):
            lbl_4.configure(fg='#0275d8')
        def key_press_lbl_4(event):
            captura = ent_campo4.get()
            lbl_4.configure(text=f'Volume {captura.replace(",", ".").upper()}')

        def cor_out_lbl_5(event):
            lbl_5.configure(fg='#222222')
        def cor_in_lbl_5(event):
            lbl_5.configure(fg='#0275d8')
        def key_press_lbl_5(event):
            captura = ent_campo5.get()
            lbl_5.configure(text=f'Corrida {captura}')

        def cor_out_lbl_6(event):
            lbl_6.configure(fg='#222222')
        def cor_in_lbl_6(event):
            lbl_6.configure(fg='#0275d8')
        def key_press_lbl_6(event):
            captura = ent_campo6.get()
            lbl_6.configure(text=f'Peso {captura} kg')

        def cor_out_lbl_8(event):
            lbl_8.configure(fg='#222222')
        def cor_in_lbl_8(event):
            lbl_8.configure(fg='#0275d8')
        def key_press_lbl_8(event):
            captura = ent_campo8.get()
            lbl_8.configure(text=f'Registro\n{captura}', justify='center')

        def impressao():
            def modelo2(): #Rolo
                l = zpl.Label(114,75)
                pos_y = 31
                pos_x = 2
                l.origin(pos_x,pos_y)
                l.write_text(
                    "VERGALHAO CA-50",
                    char_height=4,
                    char_width=4,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 4
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"{desc1} mm ROLO (S)",
                    char_height=4,
                    char_width=4,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 4
                l.origin(pos_x,pos_y)
                l.write_text(
                    "//// CABECA ////",
                    char_height=4,
                    char_width=4,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 6
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"Codigo {codigo}",
                    char_height=4,
                    char_width=3,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 4
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"Volume {volume}",
                    char_height=4,
                    char_width=3,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()


                pos_y += 4
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"Corrida {corrida}",
                    char_height=4,
                    char_width=3,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 4
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"Peso {peso} kg",
                    char_height=4,
                    char_width=3,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 5
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"{abnt}",
                    char_height=2,
                    char_width=2,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 2
                pos_x +=10
                l.origin(pos_x,pos_y+5)
                l.write_text(
                    "Registro",
                    char_height=2,
                    char_width=2,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                l.origin(pos_x-2,pos_y+7)
                l.write_text(
                    f"{inmetro}",
                    char_height=3,
                    char_width=2,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()


                l.origin(42, 17)
                l.write_barcode(height=60, barcode_type='C', check_digit='Y', orientation='R')
                l.write_text(f'{codigo}{volume}{corrida}{peso}')
                l.endorigin()

                pos_y += 8
                l.origin(12, pos_y+10)
                l.write_barcode(height=80, barcode_type='C', check_digit='Y', orientation='N')
                l.write_text(f'{volume}')
                l.endorigin()

                #print(height)

                #print(l.dumpZPL())
                pg1 = l.dumpZPL()
                pg2 = pg1.replace('//// CABECA ////', '//// CAUDA ////')
                uniao_paginas = pg1+pg2
                #print(uniao_paginas)
                
                #l.preview()

                printer_name = win32print.GetDefaultPrinter ()
                #
                # raw_data could equally be raw PCL/PS read from
                #  some print-to-file operation
                #
                if sys.version_info >= (3,):
                    #raw_data = bytes ("^XA^FO24,372^A0N,48,48^FB720,1,0,L,0^FDVERGALHAO CA-50^FS^FO24,420^A0N,48,48^FB720,1,0,L,0^FD8,00 MM ROLO^FS^FO24,468^A0N,48,48^FB720,1,0,L,0^FD//// CAUDA ////^FS^FO24,540^A0N,48,36^FB720,1,0,L,0^FDCodigo V1020B8.0^FS^FO24,588^A0N,48,36^FB720,1,0,L,0^FDVolume LHC14583^FS^FO24,636^A0N,48,36^FB720,1,0,L,0^FDCorrida 31244^FS^FO24,684^A0N,48,36^FB720,1,0,L,0^FDPeso 2,007 TON^FS^FO144,828^A0N,36,24^FB720,1,0,L,0^FDRegistro^FS^FO120,864^A0N,36,24^FB720,1,0,L,0^FD005713/2015^FS^FO504,204^BCB,60,N,N,Y,N^FDV1020B8.0 LHC1458331244 2,007^FS^FO144,996^BCN,80,S,N,Y,N^FDLHC14583^FS^XZ", "utf-8")
                    raw_data = bytes (uniao_paginas, "utf-8")
                else:
                    raw_data = "This is a test"
                hPrinter = win32print.OpenPrinter (printer_name)
                try:
                    hJob = win32print.StartDocPrinter (hPrinter, 1, ("test of raw data", None, "RAW"))
                    try:
                        win32print.StartPagePrinter (hPrinter)
                        win32print.WritePrinter (hPrinter, raw_data)
                        win32print.EndPagePrinter (hPrinter)
                    finally:
                        win32print.EndDocPrinter (hPrinter)
                finally:
                    win32print.ClosePrinter (hPrinter)

            def limpa_campos():
                ent_campo1.delete(0,END)
                ent_campo3.delete(0,END)
                ent_campo4.delete(0,END)
                ent_campo5.delete(0,END)
                ent_campo6.delete(0,END)
                ent_campo8.configure(state='normal')
                ent_campo8.delete(0, END)
                ent_campo8.configure(state='readonly')

            tipo_produto = 'ROLO'
            desc1 = ent_campo1.get().replace('.', ',')
            codigo = ent_campo3.get().upper().replace(',', '.')
            volume = ent_campo4.get().upper()
            corrida = ent_campo5.get()
            peso = ent_campo6.get()
            inmetro = ent_campo8.get()
            usuario = usuario_logado[2]
            if desc1=='' or codigo =='' or volume ==''or corrida ==''or peso ==''or inmetro =='':
                messagebox.showwarning('+Etiqueta:', 'Todos os campos devem ser preenchidos.', parent=root)
            else:
                try:
                    cursor.execute("INSERT INTO etiquetas_avulsas (\
                        data,\
                        hora,\
                        usuario,\
                        tipo_produto,\
                        desc1,\
                        codigo,\
                        volume,\
                        corrida,\
                        peso,\
                        abnt,\
                        inmetro)\
                        values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", (data, hora, usuario, tipo_produto, desc1, codigo, volume, corrida, peso, abnt, inmetro))
                    db.commit()
                except Exception as e:
                    messagebox.showerror('+Etiqueta:', 'Erro ao gravar as informações na tabela (etiquetas_avulsas).', parent=root)
                    print(e)
                    return False
                imp = threading.Thread(target=modelo2)
                imp.start()
                limpa_campos()
                ent_campo3.focus_force()

        fr3_1 = Frame(fr3, bg='#ffffff')
        fr3_1.pack(side=LEFT, fill=X, expand=True)
        fr3_2 = Frame(fr3, bg='#ffffff')
        fr3_2.pack(side=RIGHT, fill=X, expand=True)            
        
        #//////FRAME3 e 3_1
        cursor.execute("SELECT norma_abnt FROM simec_etiquetas.norma_abnt")
        abnt = cursor.fetchone()[0]
        image_avulsas = Image.open('img\\avulsas.png')
        resize_avulsas = image_avulsas.resize((290, 114))
        nova_image_avulsas = ImageTk.PhotoImage(resize_avulsas)
        lbl = Label(fr3_1, image=nova_image_avulsas, bg='#ffffff')
        lbl.grid(row=0, column=1)
        lbl = Label(fr3_1, text="VERGALHAO CA-50", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=1, column=1, sticky='W', padx=10)
        lbl_1 = Label(fr3_1, text="0,00 mm ROLO (S)", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_1.grid(row=2, column=1, sticky='W', padx=10)    
        lbl_2 = Label(fr3_1, text="//// CABECA ////", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_2.grid(row=3, column=1, sticky='W', padx=10)    
        lbl_3 = Label(fr3_1, text="Codigo VRCA00000", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_3.grid(row=4, column=1, sticky='W', padx=10, pady=(6,0))    
        lbl_4 = Label(fr3_1, text="Volume LIA00000", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_4.grid(row=5, column=1, sticky='W', padx=10)    
        lbl_5 = Label(fr3_1, text="Corrida 00000", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_5.grid(row=6, column=1, sticky='W', padx=10)    
        lbl_6 = Label(fr3_1, text="Peso 0000 kg", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_6.grid(row=7, column=1, sticky='W', padx=10)    
        lbl_7 = Label(fr3_1, text=abnt, font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl_7.grid(row=8, column=1, sticky='W', padx=10)    
        lbl_8 = Label(fr3_1, text="Registro\n000000/0000", font=fonte_padrao_bold, fg='#222222', bg='#ffffff', justify='center')
        lbl_8.grid(row=9, column=1, sticky='W', padx=(50,0), pady=(5,0))    
        
        fr3_1.grid_columnconfigure(0, weight=1)
        
        #//////FRAME3 e 3_2
        lbl = Label(fr3_2, text="Desc.1:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=1, column=2)
        ent_campo1 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo1.grid(row=1, column=3, pady=(10,8))
        ent_campo1.bind('<KeyRelease>',key_press_lbl_1)
        ent_campo1.bind("<FocusIn>", cor_in_lbl_1)
        ent_campo1.bind("<FocusOut>", cor_out_lbl_1)


        lbl = Label(fr3_2, text="Codigo:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=3, column=2)
        ent_campo3 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo3.grid(row=3, column=3, pady=(8))
        ent_campo3.bind('<KeyRelease>',key_press_lbl_3)
        ent_campo3.bind("<FocusIn>", cor_in_lbl_3)
        ent_campo3.bind("<FocusOut>", cor_out_lbl_3)

        lbl = Label(fr3_2, text="Volume:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=4, column=2)
        ent_campo4 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo4.grid(row=4, column=3, pady=(8))
        ent_campo4.bind('<KeyRelease>',key_press_lbl_4)
        ent_campo4.bind("<FocusIn>", cor_in_lbl_4)
        ent_campo4.bind("<FocusOut>", cor_out_lbl_4)

        lbl = Label(fr3_2, text="Corrida:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=5, column=2)
        ent_campo5 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo5.grid(row=5, column=3, pady=(8))
        ent_campo5.bind('<KeyRelease>',key_press_lbl_5)
        ent_campo5.bind("<FocusIn>", cor_in_lbl_5)
        ent_campo5.bind("<FocusOut>", cor_out_lbl_5)

        lbl = Label(fr3_2, text="Peso:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=6, column=2)
        ent_campo6 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo6.grid(row=6, column=3, pady=(8))
        ent_campo6.bind('<KeyRelease>',key_press_lbl_6)
        ent_campo6.bind("<FocusIn>", cor_in_lbl_6)
        ent_campo6.bind("<FocusOut>", cor_out_lbl_6)

        lbl = Label(fr3_2, text="Inmetro:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=7, column=2)
        ent_campo8 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo8.grid(row=7, column=3, pady=(8))
        ent_campo8.bind('<KeyRelease>',key_press_lbl_8)
        ent_campo8.bind("<FocusIn>", cor_in_lbl_8)
        ent_campo8.bind("<FocusOut>", cor_out_lbl_8)
        ent_campo8.configure(state='readonly')        

        fr3_2.grid_columnconfigure(4, weight=1)
        
        #//////FRAME5
        fr5.grid_columnconfigure(0, weight=1)
        fr5.grid_columnconfigure(6, weight=1)
        global bt_imp
        bt_imp = customtkinter.CTkButton(fr5, text='Imprimir (F1)', **estilo_botao_padrao_etiqueta, command=impressao)
        bt_imp.grid(row=1, column=1)
        root.bind('<F1>', lambda event: impressao())        
        mainloop()    

    def fiomaquina():
        for widget in fr3.winfo_children():
            widget.destroy()
        for widget in fr5.winfo_children():
            widget.destroy()

        def cor_out_lbl_1(event):
            lbl_1.configure(fg='#222222')
        def cor_in_lbl_1(event):
            lbl_1.configure(fg='#0275d8')
        def key_press_lbl_1(event):
            captura = ent_campo1.get()
            lbl_1.configure(text=f'{captura.replace(".", ",")} mm BARRA')

        def cor_out_lbl_2(event):
            lbl_2.configure(fg='#222222')
        def cor_in_lbl_2(event):
            lbl_2.configure(fg='#0275d8')
        def key_press_lbl_2(event):
            captura = ent_campo2.get()
            lbl_2.configure(text=f'ACO {captura.upper()}')

        def cor_out_lbl_3(event):
            lbl_3.configure(fg='#222222')
        def cor_in_lbl_3(event):
            lbl_3.configure(fg='#0275d8')
        def key_press_lbl_3(event):
            captura = ent_campo3.get()
            lbl_3.configure(text=f'Codigo {captura.replace(",", ".").upper()}')

        def cor_out_lbl_4(event):
            lbl_4.configure(fg='#222222')
        def cor_in_lbl_4(event):
            lbl_4.configure(fg='#0275d8')
        def key_press_lbl_4(event):
            captura = ent_campo4.get()
            lbl_4.configure(text=f'Volume {captura.replace(",", ".").upper()}')

        def cor_out_lbl_5(event):
            lbl_5.configure(fg='#222222')
        def cor_in_lbl_5(event):
            lbl_5.configure(fg='#0275d8')
        def key_press_lbl_5(event):
            captura = ent_campo5.get()
            lbl_5.configure(text=f'Corrida {captura}')

        def cor_out_lbl_6(event):
            lbl_6.configure(fg='#222222')
        def cor_in_lbl_6(event):
            lbl_6.configure(fg='#0275d8')
        def key_press_lbl_6(event):
            captura = ent_campo6.get()
            lbl_6.configure(text=f'Peso {captura} kg')

        def impressao():
            def modelo3(): #FM
                l = zpl.Label(114,75)
                pos_y = 30
                pos_x = 2
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"{desc1} mm",
                    char_height=4,
                    char_width=4,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()
                #6.50 MM|AÇO 1060
                pos_y += 4
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"ACO {desc2}",
                    char_height=4,
                    char_width=4,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 4
                l.origin(pos_x,pos_y)
                l.write_text(
                    "//// CABECA ////",
                    char_height=4,
                    char_width=4,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 6
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"Codigo {codigo}",
                    char_height=4,
                    char_width=3,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 4
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"Volume {volume}",
                    char_height=4,
                    char_width=3,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()


                pos_y += 4
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"Corrida {corrida}",
                    char_height=4,
                    char_width=3,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 4
                l.origin(pos_x,pos_y)
                l.write_text(
                    f"Peso {peso} kg",
                    char_height=4,
                    char_width=3,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 5
                l.origin(pos_x,pos_y)
                l.write_text(
                    "",
                    char_height=2,
                    char_width=2,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                pos_y += 3
                pos_x +=10
                l.origin(pos_x,pos_y+5)
                l.write_text(
                    "",
                    char_height=2,
                    char_width=2,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                l.origin(pos_x-2,pos_y+7)
                l.write_text(
                    f"",
                    char_height=3,
                    char_width=2,
                    line_width=60,
                    justification='L',
                    )
                l.endorigin()

                l.origin(42, 17)
                l.write_barcode(height=60, barcode_type='C', check_digit='Y', orientation='R')
                l.write_text(f'{codigo}{volume}{corrida}{peso}')
                l.endorigin()

                pos_y += 8
                l.origin(12, pos_y+10)
                l.write_barcode(height=80, barcode_type='C', check_digit='Y', orientation='N')
                l.write_text(f'{volume}')
                l.endorigin()

                #l.preview()
                #print(height)

                pg1 = l.dumpZPL()
                pg2 = pg1.replace('//// CABECA ////', '//// CAUDA ////')
                uniao_paginas = pg1+pg2

                printer_name = win32print.GetDefaultPrinter ()
                #
                # raw_data could equally be raw PCL/PS read from
                #  some print-to-file operation
                #
                if sys.version_info >= (3,):
                    #raw_data = bytes ("^XA^FO24,372^A0N,48,48^FB720,1,0,L,0^FDVERGALHAO CA-50^FS^FO24,420^A0N,48,48^FB720,1,0,L,0^FD8,00 MM ROLO^FS^FO24,468^A0N,48,48^FB720,1,0,L,0^FD//// CAUDA ////^FS^FO24,540^A0N,48,36^FB720,1,0,L,0^FDCodigo V1020B8.0^FS^FO24,588^A0N,48,36^FB720,1,0,L,0^FDVolume LHC14583^FS^FO24,636^A0N,48,36^FB720,1,0,L,0^FDCorrida 31244^FS^FO24,684^A0N,48,36^FB720,1,0,L,0^FDPeso 2,007 TON^FS^FO144,828^A0N,36,24^FB720,1,0,L,0^FDRegistro^FS^FO120,864^A0N,36,24^FB720,1,0,L,0^FD005713/2015^FS^FO504,204^BCB,60,N,N,Y,N^FDV1020B8.0 LHC1458331244 2,007^FS^FO144,996^BCN,80,S,N,Y,N^FDLHC14583^FS^XZ", "utf-8")
                    raw_data = bytes (uniao_paginas, "utf-8")
                    #print(raw_data) 
                else:
                    raw_data = "This is a test"
                hPrinter = win32print.OpenPrinter (printer_name)
                try:
                    hJob = win32print.StartDocPrinter (hPrinter, 1, ("test of raw data", None, "RAW"))
                    try:
                        win32print.StartPagePrinter (hPrinter)
                        win32print.WritePrinter (hPrinter, raw_data)
                        win32print.EndPagePrinter (hPrinter)
                    finally:
                        win32print.EndDocPrinter (hPrinter)
                finally:
                    win32print.ClosePrinter (hPrinter)                    

            def limpa_campos():
                ent_campo1.delete(0,END)
                ent_campo2.delete(0,END)
                ent_campo3.delete(0,END)
                ent_campo4.delete(0,END)
                ent_campo5.delete(0,END)
                ent_campo6.delete(0,END)

            tipo_produto = 'FIO MAQUINA'
            desc1 = ent_campo1.get().replace('.', ',')
            desc2 = ent_campo2.get().upper()
            codigo = ent_campo3.get().upper().replace(',', '.')
            volume = ent_campo4.get().upper()
            corrida = ent_campo5.get()
            peso = ent_campo6.get()
            usuario = usuario_logado[2]
            if desc1=='' or desc2 =='' or codigo =='' or volume ==''or corrida ==''or peso =='':
                messagebox.showwarning('+Etiqueta:', 'Todos os campos devem ser preenchidos.', parent=root)
            else:
                try:
                    cursor.execute("INSERT INTO etiquetas_avulsas (\
                        data,\
                        hora,\
                        usuario,\
                        tipo_produto,\
                        desc1,\
                        desc2,\
                        codigo,\
                        volume,\
                        corrida,\
                        peso)\
                        values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", (data, hora, usuario, tipo_produto, desc1, desc2, codigo, volume, corrida, peso))
                    db.commit()
                except Exception as e:
                    messagebox.showerror('+Etiqueta:', 'Erro ao gravar as informações na tabela (etiquetas_avulsas).', parent=root)
                    print(e)
                    return False
                imp = threading.Thread(target=modelo3)
                imp.start()
                limpa_campos()
                ent_campo2.focus_force()

        fr3_1 = Frame(fr3, bg='#ffffff')
        fr3_1.pack(side=LEFT, fill=X, expand=True)
        fr3_2 = Frame(fr3, bg='#ffffff')
        fr3_2.pack(side=RIGHT, fill=X, expand=True)            
        
        #//////FRAME3 e 3_1
        image_avulsas = Image.open('img\\avulsas2.png')
        resize_avulsas = image_avulsas.resize((290, 114))
        nova_image_avulsas = ImageTk.PhotoImage(resize_avulsas)
        lbl = Label(fr3_1, image=nova_image_avulsas, bg='#ffffff')
        lbl.grid(row=0, column=1)
        lbl_1 = Label(fr3_1, text="0,00 mm", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_1.grid(row=2, column=1, sticky='W', padx=10)    
        lbl_2 = Label(fr3_1, text="ACO 0000", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_2.grid(row=3, column=1, sticky='W', padx=10)    
        lbl_2_1 = Label(fr3_1, text="//// CABECA ////", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_2_1.grid(row=4, column=1, sticky='W', padx=10)    
        lbl_3 = Label(fr3_1, text="Codigo FM00000", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_3.grid(row=5, column=1, sticky='W', padx=10, pady=(6,0))    
        lbl_4 = Label(fr3_1, text="Volume LIA00000", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_4.grid(row=6, column=1, sticky='W', padx=10)    
        lbl_5 = Label(fr3_1, text="Corrida 00000", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_5.grid(row=7, column=1, sticky='W', padx=10)    
        lbl_6 = Label(fr3_1, text="Peso 0000 kg", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl_6.grid(row=8, column=1, sticky='W', padx=10)    

        fr3_1.grid_columnconfigure(0, weight=1)
        
        #//////FRAME3 e 3_2
        lbl = Label(fr3_2, text="Desc.1:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=1, column=2)
        ent_campo1 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo1.grid(row=1, column=3, pady=(10,8))
        ent_campo1.bind('<KeyRelease>',key_press_lbl_1)
        ent_campo1.bind("<FocusIn>", cor_in_lbl_1)
        ent_campo1.bind("<FocusOut>", cor_out_lbl_1)

        lbl = Label(fr3_2, text="Desc.2:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=2, column=2)
        ent_campo2 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo2.grid(row=2, column=3, pady=(8))
        ent_campo2.bind('<KeyRelease>',key_press_lbl_2)
        ent_campo2.bind("<FocusIn>", cor_in_lbl_2)
        ent_campo2.bind("<FocusOut>", cor_out_lbl_2)

        lbl = Label(fr3_2, text="Codigo:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=3, column=2)
        ent_campo3 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo3.grid(row=3, column=3, pady=(8))
        ent_campo3.bind('<KeyRelease>',key_press_lbl_3)
        ent_campo3.bind("<FocusIn>", cor_in_lbl_3)
        ent_campo3.bind("<FocusOut>", cor_out_lbl_3)

        lbl = Label(fr3_2, text="Volume:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=4, column=2)
        ent_campo4 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo4.grid(row=4, column=3, pady=(8))
        ent_campo4.bind('<KeyRelease>',key_press_lbl_4)
        ent_campo4.bind("<FocusIn>", cor_in_lbl_4)
        ent_campo4.bind("<FocusOut>", cor_out_lbl_4)

        lbl = Label(fr3_2, text="Corrida:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=5, column=2)
        ent_campo5 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo5.grid(row=5, column=3, pady=(8))
        ent_campo5.bind('<KeyRelease>',key_press_lbl_5)
        ent_campo5.bind("<FocusIn>", cor_in_lbl_5)
        ent_campo5.bind("<FocusOut>", cor_out_lbl_5)

        lbl = Label(fr3_2, text="Peso:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=6, column=2)
        ent_campo6 = customtkinter.CTkEntry(fr3_2, **estilo_entry_padrao, width=220)
        ent_campo6.grid(row=6, column=3, pady=(8))
        ent_campo6.bind('<KeyRelease>',key_press_lbl_6)
        ent_campo6.bind("<FocusIn>", cor_in_lbl_6)
        ent_campo6.bind("<FocusOut>", cor_out_lbl_6)

        fr3_2.grid_columnconfigure(4, weight=1)
        
        #//////FRAME5
        fr5.grid_columnconfigure(0, weight=1)
        fr5.grid_columnconfigure(6, weight=1)
        global bt_imp
        bt_imp = customtkinter.CTkButton(fr5, text='Imprimir (F1)', **estilo_botao_padrao_etiqueta, command=impressao)
        bt_imp.grid(row=1, column=1)
        root.bind('<F1>', lambda event: impressao())        
        mainloop()    

    def historico():

        def atualizar_lista_impressao():
            db.cmd_reset_connection()
            print(data)
            tree_hist.delete(*tree_hist.get_children())
            cursor.execute("SELECT * FROM simec_etiquetas.etiquetas_avulsas order by id DESC")
            cont = 0
            for row in cursor:
                if cont % 2 == 0:
                    tree_hist.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[4], row[3], row[8], row[9], row[10], row[1], row[2]),
                                            tags=('par',))
                else:
                    tree_hist.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[4], row[3], row[8], row[9], row[10], row[1], row[2]),
                                            tags=('impar',))
                cont += 1

        for widget in fr3.winfo_children():
            widget.destroy()
        for widget in fr5.winfo_children():
            widget.destroy()

        style = ttk.Style()
        # style.theme_use('default')
        style.configure('Treeview',
                        background='#ffffff',
                        rowheight=24,
                        fieldbackground='#ffffff',
                        font=fonte_padrao)
        style.configure("Treeview.Heading",
                        foreground='#1d366c',
                        background="#ffffff",
                        font=fonte_padrao_bold)
        style.map('Treeview', background=[('selected', '#1D366C')])

        tree_hist = ttk.Treeview(fr3, selectmode='browse')
        vsb = ttk.Scrollbar(fr3, orient="vertical", command=tree_hist.yview)
        vsb.pack(side=RIGHT, fill='y')
        tree_hist.configure(yscrollcommand=vsb.set)
        vsbx = ttk.Scrollbar(fr3, orient="horizontal", command=tree_hist.xview)
        vsbx.pack(side=BOTTOM, fill='x')
        tree_hist.configure(xscrollcommand=vsbx.set)
        tree_hist.pack(side=LEFT, fill=BOTH, expand=True, anchor='n')
        tree_hist["columns"] = ("1", "2", "3", "4", "5", "6", "7", "8")
        tree_hist['show'] = 'headings'
        tree_hist.column("#1", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#2", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#3", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#4", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#5", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#6", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#7", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#8", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.heading("1", text="Registro DB")
        tree_hist.heading("2", text="Produto")
        tree_hist.heading("3", text="Usuário")
        tree_hist.heading("4", text="Volume")
        tree_hist.heading("5", text="Corrida")
        tree_hist.heading("6", text="Peso")
        tree_hist.heading("7", text="Data")
        tree_hist.heading("8", text="Hora")
        tree_hist.tag_configure('par', background='#e9e9e9')
        tree_hist.tag_configure('impar', background='#ffffff')
        #tree_hist.bind("<Double-1>", duploclique_tree_hist)
        fr3.grid_columnconfigure(0, weight=1)
        fr3.grid_columnconfigure(3, weight=1)
        atualizar_lista_impressao()     

    def menu_opcoes(value):
        if value == '   Barra/Endireitado   ':
            barra_endireitado()
        elif value == '    Rolo    ':
            rolo()
        elif value == '    Fio Máquina    ':
            fiomaquina()
        elif value == '    Histórico    ':
            historico()

    fr1 = Frame(frame4, bg='#ffffff')
    fr1.pack(side=TOP, fill=X)
    fr2 = Frame(frame4, bg='#222222')  # linha
    fr2.pack(side=TOP, fill=X, pady=5)
    fr2_1 = Frame(frame4, bg='#ffffff')
    fr2_1.pack(side=TOP, fill=X)
    fr3 = Frame(frame4, bg='#ffffff')
    fr3.pack(side=TOP, fill=BOTH, expand=True)
    fr4 = Frame(frame4, bg='#222222')  # linha
    fr4.pack(side=TOP, fill=X, pady=5)
    fr5 = Frame(frame4, bg='#ffffff')
    fr5.pack(side=TOP, fill=X, pady=(0,20))

    #//////FRAME1
    image_invent = Image.open('img\\avulsa2.png')
    resize_invent = image_invent.resize((50, 50))
    nova_image_invent = ImageTk.PhotoImage(resize_invent)
    lbl = Label(fr1, text=" Etiquetas Avulsas:", image=nova_image_invent, compound="left", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
    lbl.grid(row=0, column=1)
    fr1.grid_columnconfigure(0, weight=1)
    fr1.grid_columnconfigure(2, weight=1)
    #//////FRAME2 LINHA
    
    #//////FRAME2_1


    #//////FRAME2_1
    button_var = customtkinter.StringVar(value="   Barra/Endireitado   ")  # set initial value
    segemented_button = customtkinter.CTkSegmentedButton(
        fr2_1,
        values=["   Barra/Endireitado   ", "    Rolo    ", "    Fio Máquina    ","    Histórico    "],
        variable=button_var, command=menu_opcoes, bg_color="#ffffff", fg_color="#dcdcdc", height=30, font=fonte_padrao_bold,
        selected_color="#222222", selected_hover_color="#363636",unselected_color="#1D366C", unselected_hover_color="#363636")
    
    segemented_button.pack(pady=10)
   
    #//////FRAME4 LINHA

    #//////FRAME5
    
    barra_endireitado()

    '''fr5.grid_columnconfigure(0, weight=1)
    fr5.grid_columnconfigure(6, weight=1)


    #//////FRAME6
    fr6.grid_columnconfigure(0, weight=1)
    fr6.grid_columnconfigure(6, weight=1)


    #//////FRAME7 LINHA
    
    #//////FRAME8

    
    #//////FRAME9 LINHA
    
    #//////FRAME10
    lbl_hist=Label(fr10, text='Histórico de Impressões: ', font=fonte_titulo_bold, bg='#ffffff', fg='#222222')
    lbl_hist.grid(row=0, column=1, padx=5, pady=(0,5))

    bt_select = customtkinter.CTkButton(fr10, text='Selecionar Item', **estilo_botao_padrao_etiqueta, command=select)
    bt_select.grid(row=0, column=2)
    
    fr10.grid_columnconfigure(3, weight=1)
    
    clique_pesq = StringVar()
    lista_pesq = ['Corrida','Data','Produto','Usuário','Volume','Remover Filtro']

    opt_pesq = customtkinter.CTkOptionMenu(fr10, variable=clique_pesq, values=lista_pesq, width=160, **estilo_optionmenu,command=opt_pesq_imp)
    opt_pesq.grid(row=0, column=4, padx=2)
    opt_pesq.set('Pesquisar por...')
    
    ent_pesquisa = customtkinter.CTkEntry(fr10, **estilo_entry_padrao, width=160)
    ent_pesquisa.grid(row=0, column=5)
    ent_pesquisa.bind("<Return>", pesquisar_imp_bind)

    image_pesq = Image.open('img\\pesq.png')
    resize_pesq = image_pesq.resize((30, 30))
    nova_image_pesq = ImageTk.PhotoImage(resize_pesq)
    btn_pesq = Button(fr10, image=nova_image_pesq, compound="top", bg='#FFFFFF', fg='#FFFFFF',
                    activebackground='#FFFFFF', borderwidth=0,
                    activeforeground="#FFFFFF", highlightthickness=4, relief=RIDGE, command=pesquisar_imp,
                    font=fonte_botoes, cursor="hand2")
    btn_pesq.grid(row=0, column=6, padx=0)

    #//////FRAME11


    #//////FRAME12
    
    bt_hist_reimp = customtkinter.CTkButton(fr12, text='Histórico de Reimpressões', **estilo_botao_padrao_vermelho, command=hist_reimp)
    bt_hist_reimp.grid(row=0, column=1, padx=10, pady=5)
    
    fr12.grid_columnconfigure(2, weight=1)

    atualizar_lista_impressao()
    setup_reimp()'''
    mainloop()

def corridas():
    global controle_loop
    controle_loop = ''
        
    for widget in frame4.winfo_children():
        widget.destroy()

    def atualizar_lista_usuarios():
        db.cmd_reset_connection()
        tree_corrida.delete(*tree_corrida.get_children())
        cursor.execute("SELECT * from corridas ORDER BY id DESC")
        cont = 0
        for row in cursor:
            if cont % 2 == 0:
                tree_corrida.insert('', 'end', text=" ",
                                        values=(
                                        row[0], row[1], row[2], row[3], row[4], row[5]),
                                        tags=('par',))
            else:
                tree_corrida.insert('', 'end', text=" ",
                                        values=(
                                        row[0], row[1], row[2], row[3], row[4], row[5]),
                                        tags=('impar',))
            cont += 1

    def salvar():
        corrida = ent_corrida.get().upper()
        qtd_pecas = ent_pecas.get()
        int_corrida= corrida.isnumeric()
        int_qtd_pecas= qtd_pecas.isnumeric()
        if corrida == '' or qtd_pecas == '':
            messagebox.showwarning('Cadastro de Corridas:', 'Todos os campos devem ser preenchidos.', parent=root)
        elif len(corrida) > 5:
            messagebox.showwarning('Cadastro de Corridas:', 'Número de caracteres inválidos.', parent=root)
        elif int_corrida == False or int_qtd_pecas == False :
            messagebox.showwarning('Cadastro de Corridas:', 'Somente números inteiros são permitidos.', parent=root)
        else:
            cursor.execute("SELECT * FROM corridas WHERE corrida=%s", (corrida,))
            verifica_corrida = cursor.fetchone()
            if verifica_corrida == None:
                try:
                    cursor.execute(
                        "INSERT INTO corridas (corrida, qtd_pecas, qtd_perdas, qtd_consumidas, encerrada) values(%s,%s,%s,%s,%s)", (corrida,qtd_pecas,0,0, 'Não'))
                    db.commit()
                except Exception as e:
                    messagebox.showerror('Cadastro de Corridas:', 'Erro de conexão com o Banco de Dados.', parent=root)
                    print(e)
                    return False
                messagebox.showinfo('Cadastro de Corridas:', 'Corrida cadastrada com sucesso.', parent=root)
                corridas()
            else:
                messagebox.showerror('Cadastro de Corridas', 'Corrida já cadastrada:', parent=root)

    def editar():
            def setup_interno():
                ent_corrida.delete(0, END)
                ent_corrida.insert(0, result2[1])
                ent_pecas.delete(0, END)
                ent_pecas.insert(0, result2[2])
                opt_encerrada.set(result2[5])

            def confirmar():
                corrida = ent_corrida.get().upper()
                pecas = ent_pecas.get()
                int_corrida= corrida.isnumeric()
                int_qtd_pecas= pecas.isnumeric()
                encerrada = clique_encerrada.get()

                if corrida == '' or pecas == '':
                    messagebox.showwarning('Cadastro de Corridas:', 'Todos os campos devem ser preenchidos.', parent=root)
                elif len(corrida) > 5 or len(corrida) < 5:
                    messagebox.showwarning('Cadastro de Corridas:', 'Número de caracteres inválidos.', parent=root)
                elif int_corrida == False or int_qtd_pecas == False :
                    messagebox.showwarning('Cadastro de Corridas:', 'Somente números inteiros são permitidos.', parent=root)                
                else:
                    try:
                        cursor.execute(
                            "UPDATE corridas SET\
                            corrida = %s,\
                            qtd_pecas = %s,\
                            encerrada = %s\
                            WHERE id = %s", (corrida,pecas,encerrada,result2[0],))
                        db.commit()
                    except Exception as e:
                        messagebox.showerror('Cadastro de Corridas:', 'Erro de conexão com o Banco de Dados.', parent=root)
                        print(e)
                        return False
                    messagebox.showinfo('Cadastro de Corridas:', 'Edição realizada com sucesso.', parent=root)
                    corridas()
            
            lista_select = tree_corrida.focus()
            if lista_select == "":
                messagebox.showwarning('Cadastro de Corridas:', 'Selecione um item na lista!', parent=root)
            else:
                valor_lista = tree_corrida.item(lista_select, "values")[0]
                try:
                    cursor.execute("SELECT * from corridas WHERE id = %s",(valor_lista,))
                    result2 = cursor.fetchone()
                    #print(result2)
                except Exception as e:
                    messagebox.showerror('Cadastro de Corridas:', 'Erro de conexão com o Banco de Dados.', parent=root)
                    print(e)
                    return False
                btn_salvar_produto.grid_remove()
                btn_editar_produto.grid_remove()

                image_confirmar = Image.open('img\\salvar.png')
                resize_confirmar = image_confirmar.resize((28, 32))
                nova_image_confirmar = ImageTk.PhotoImage(resize_confirmar)
                btn_confirmar_produto =Button(fr8, image=nova_image_confirmar, text="Confirmar", compound="top",
                                    font=fonte_padrao_bold, bg='#ffffff', fg='#222222', command=confirmar,
                                    borderwidth=0, relief=RIDGE, activebackground="#ffffff", activeforeground="#1d366c", cursor="hand2")
                btn_confirmar_produto.photo = nova_image_confirmar
                btn_confirmar_produto.grid(row=0, column=1, padx=4)

                image_cancelar = Image.open('img\\cancelar.png')
                resize_cancelar = image_cancelar.resize((28, 32))
                nova_image_cancelar = ImageTk.PhotoImage(resize_cancelar)
                btn_cancelar_produto =Button(fr8, image=nova_image_cancelar, text="Cancelar", compound="top",
                                    font=fonte_padrao_bold, bg='#ffffff', fg='#222222', command=corridas,
                                    borderwidth=0, relief=RIDGE, activebackground="#ffffff", activeforeground="#1d366c", cursor="hand2")
                btn_cancelar_produto.photo = nova_image_cancelar
                btn_cancelar_produto.grid(row=0, column=2, padx=4)

                lbl_corrida.configure(text='Editando a corrida de nº', fg='#8B0000')

                lbl_pecas=Label(fr3, text='Encerrada?:', font=fonte_padrao, bg='#ffffff', fg='#000000')
                lbl_pecas.grid(row=4, column=1, sticky="w")
                clique_encerrada = StringVar()
                lista_opcoes = ['Sim', 'Não']
                opt_encerrada = customtkinter.CTkOptionMenu(fr3, variable=clique_encerrada, values=lista_opcoes, width=296, **estilo_optionmenu )
                opt_encerrada.grid(row=5, column=1)


                setup_interno()

    def pesquisar_corridas_bind(event):
        pesquisar_corridas()

    def pesquisar_corridas():
        pesq = ent_pesquisa.get()
        if pesq != '':
            db.cmd_reset_connection()
            tree_corrida.delete(*tree_corrida.get_children())
            cursor.execute("SELECT * FROM corridas where corrida like concat('%',%s,'%')",(pesq,))
            cont = 0
            for row in cursor:
                if cont % 2 == 0:
                    tree_corrida.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[1], row[2], row[3], row[4], row[5]),
                                            tags=('par',))
                else:
                    tree_corrida.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[1], row[2], row[3], row[4], row[5]),
                                            tags=('impar',))
                cont += 1
    
    def desfaz_pesquisa():
        db.cmd_reset_connection()
        tree_corrida.delete(*tree_corrida.get_children())
        cursor.execute("SELECT * FROM corridas ")
        cont = 0
        for row in cursor:
            if cont % 2 == 0:
                tree_corrida.insert('', 'end', text=" ",
                                        values=(
                                        row[0], row[1], row[2], row[3], row[4], row[5]),
                                        tags=('par',))
            else:
                tree_corrida.insert('', 'end', text=" ",
                                        values=(
                                        row[0], row[1], row[2], row[3], row[4], row[5]),
                                        tags=('impar',))
            cont += 1        

    fr1 = Frame(frame4, bg='#ffffff')
    fr1.pack(side=TOP, fill=X)
    fr2 = Frame(frame4, bg='#222222')  # linha
    fr2.pack(side=TOP, fill=X, pady=5)
    fr3 = Frame(frame4, bg='#ffffff')
    fr3.pack(side=TOP, fill=X)
    fr4 = Frame(frame4, bg='#ffffff')
    fr4.pack(side=TOP, fill=X)
    fr5 = Frame(frame4, bg='#222222')  # linha
    fr5.pack(side=TOP, fill=X, pady=5)
    fr6 = Frame(frame4, bg='#ffffff')
    fr6.pack(side=TOP, fill=BOTH, expand=TRUE)
    fr7 = Frame(frame4, bg='#222222')  # linha
    fr7.pack(side=TOP, fill=X, pady=5)
    fr8 = Frame(frame4, bg='#ffffff')
    fr8.pack(side=TOP, fill=X)   

    #//////FRAME1
    image_invent = Image.open('img\\corrida2.png')
    resize_invent = image_invent.resize((50, 50))
    nova_image_invent = ImageTk.PhotoImage(resize_invent)
    lbl = Label(fr1, text=" Cadastro | Corridas", image=nova_image_invent, compound="left", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
    lbl.grid(row=0, column=0, padx=(20,0))

    ent_pesquisa = customtkinter.CTkEntry(fr1, **estilo_entry_padrao, width=160, placeholder_text='Digite a corrida..')
    ent_pesquisa.grid(row=0, column=2)
    ent_pesquisa.bind("<Return>", pesquisar_corridas_bind)

    image_pesq = Image.open('img\\pesq.png')
    resize_pesq = image_pesq.resize((30, 30))
    nova_image_pesq = ImageTk.PhotoImage(resize_pesq)
    btn_pesq = Button(fr1, image=nova_image_pesq, compound="top", bg='#FFFFFF', fg='#FFFFFF',
                    activebackground='#FFFFFF', borderwidth=0,
                    activeforeground="#FFFFFF", highlightthickness=4, relief=RIDGE, command=pesquisar_corridas,
                    font=fonte_botoes, cursor="hand2")
    btn_pesq.grid(row=0, column=3, padx=0)

    image_desfazer = Image.open('img\\desfazer.png')
    resize_desfazer = image_desfazer.resize((25, 25))
    nova_image_desfazer = ImageTk.PhotoImage(resize_desfazer)
    btn_desfazer = Button(fr1, image=nova_image_desfazer, compound="top", bg='#FFFFFF', fg='#FFFFFF',
                    activebackground='#FFFFFF', borderwidth=0,
                    activeforeground="#FFFFFF", highlightthickness=4, relief=RIDGE, command=desfaz_pesquisa,
                    font=fonte_botoes, cursor="hand2")
    btn_desfazer.grid(row=0, column=4, padx=(0,20))

    fr1.grid_columnconfigure(1, weight=1)

    #//////FRAME2 LINHA
    
    #//////FRAME3
    lbl_corrida=Label(fr3, text='Nº da Corrida:', font=fonte_padrao, bg='#ffffff', fg='#000000')
    lbl_corrida.grid(row=0, column=1, sticky="w")
    ent_corrida = customtkinter.CTkEntry(fr3, **estilo_entry_padrao, width=300)
    ent_corrida.grid(row=1, column=1, pady=3)        
    ent_corrida.focus_force()
    
    lbl_pecas=Label(fr3, text='Qtd de Peças:', font=fonte_padrao, bg='#ffffff', fg='#000000')
    lbl_pecas.grid(row=2, column=1, sticky="w")
    ent_pecas = customtkinter.CTkEntry(fr3, **estilo_entry_padrao, width=300)
    ent_pecas.grid(row=3, column=1, pady=3)        

    fr3.grid_columnconfigure(0, weight=1)
    fr3.grid_columnconfigure(6, weight=1)    


    #fr4.grid_columnconfigure(0, weight=1)
    fr4.grid_columnconfigure(6, weight=1)    

    #//FRAME5 LINHA

    #//FRAME6
    style = ttk.Style()
    # style.theme_use('default')
    style.configure('Treeview',
                    background='#ffffff',
                    rowheight=24,
                    fieldbackground='#ffffff',
                    font=fonte_padrao)
    style.configure("Treeview.Heading",
                    foreground='#000000',
                    background="#ffffff",
                    font=fonte_padrao)
    style.map('Treeview', background=[('selected', '#1d366c')])

    tree_corrida = ttk.Treeview(fr6, selectmode='browse')
    vsb = ttk.Scrollbar(fr6, orient="vertical", command=tree_corrida.yview)
    vsb.pack(side=RIGHT, fill='y')
    #tree_corrida.configure(yscrollcommand=vsb.set)
    vsbx = ttk.Scrollbar(fr6, orient="horizontal", command=tree_corrida.xview)
    vsbx.pack(side=BOTTOM, fill='x')
    tree_corrida.configure(xscrollcommand=vsbx.set)
    tree_corrida.pack(side=LEFT, fill=BOTH, expand=True, anchor='n')
    tree_corrida["columns"] = ("1", "2", "3", "4", "5", "6")
    tree_corrida['show'] = 'headings'
    tree_corrida.column("#1", anchor='c', minwidth=50, width=100, stretch = True)
    tree_corrida.column("#2", anchor='c', minwidth=50, width=100, stretch = True)
    tree_corrida.column("#3", anchor='c', minwidth=50, width=100, stretch = True)
    tree_corrida.column("#4", anchor='c', minwidth=50, width=100, stretch = True)
    tree_corrida.column("#5", anchor='c', minwidth=50, width=100, stretch = True)
    tree_corrida.column("#6", anchor='c', minwidth=50, width=100, stretch = True)    
    tree_corrida.heading("1", text="ID")
    tree_corrida.heading("2", text="Corrida")
    tree_corrida.heading("3", text="Qtd Peças")
    tree_corrida.heading("4", text="Qtd Perdas")
    tree_corrida.heading("5", text="Qtd Consumidas")
    tree_corrida.heading("6", text="Encerrada?")    
    tree_corrida.tag_configure('par', background='#e9e9e9')
    tree_corrida.tag_configure('impar', background='#ffffff')
    
    #//FRAME7 LINHA

    #//FRAME8
    image_salvar = Image.open('img\\salvar.png')
    resize_salvar = image_salvar.resize((28, 32))
    nova_image_salvar = ImageTk.PhotoImage(resize_salvar)
    btn_salvar_produto =Button(fr8, image=nova_image_salvar, text="Salvar", compound="top",
                        font=fonte_padrao_bold, bg='#ffffff', fg='#222222', command=salvar,
                        borderwidth=0, relief=RIDGE, activebackground="#ffffff", activeforeground="#1d366c", cursor="hand2")
    btn_salvar_produto.photo = nova_image_salvar
    btn_salvar_produto.grid(row=0, column=1, padx=4)

    image_editar = Image.open('img\\editar.png')
    resize_editar = image_editar.resize((28, 32))
    nova_image_editar = ImageTk.PhotoImage(resize_editar)
    btn_editar_produto =Button(fr8, image=nova_image_editar, text="Editar", compound="top",
                        font=fonte_padrao_bold, bg='#ffffff', fg='#222222', command=editar,
                        borderwidth=0, relief=RIDGE, activebackground="#ffffff", activeforeground="#1d366c", cursor="hand2")
    btn_editar_produto.photo = nova_image_editar
    btn_editar_produto.grid(row=0, column=2, padx=4, pady=10)

    fr8.grid_columnconfigure(0, weight=1)
    fr8.grid_columnconfigure(3, weight=1)

    atualizar_lista_usuarios()
    mainloop()

def reimp():
    global controle_loop
    controle_loop = ''

    def hist_reimp():
        for widget in frame4.winfo_children():
            widget.destroy()

        def atualizar_lista_reimpressao():
            db.cmd_reset_connection()
            tree_hist.delete(*tree_hist.get_children())
            cursor.execute("SELECT\
                A.id,\
                A.data,\
                A.hora,\
                A.usuario,\
                B.nome,\
                A.prod_descricao,\
                A.prod_codigo,\
                A.corrida,\
                A.lote,\
                A.peso FROM producao_reimp A\
                inner join prod_tipo B on B.id = A.id_prod_tipo\
                ORDER BY A.id DESC")
            cont = 0
            for row in cursor:
                if cont % 2 == 0:
                    tree_hist.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                            tags=('par',))
                else:
                    tree_hist.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                            tags=('impar',))
                cont += 1

        def opt_pesq_imp(event):
            if clique_pesq.get() == 'Remover Filtro':
                ent_pesquisa.delete(0,END)
                opt_pesq.set('Pesquisar por...')
                atualizar_lista_reimpressao()
        
        def pesquisar_reimp_bind(event):
            pesquisar_reimp()

        def pesquisar_reimp():
            opcao = clique_pesq.get()
            pesq = ent_pesquisa.get()
            if opcao == 'Corrida':
                db.cmd_reset_connection()
                tree_hist.delete(*tree_hist.get_children())
                cursor.execute("SELECT * FROM producao_reimp where corrida like concat('%',%s,'%')",(pesq,))
                cont = 0
                for row in cursor:
                    if cont % 2 == 0:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('par',))
                    else:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('impar',))
                    cont += 1

            elif opcao == 'Data':
                db.cmd_reset_connection()
                tree_hist.delete(*tree_hist.get_children())
                cursor.execute("SELECT * FROM producao_reimp where data like concat('%',%s,'%')",(pesq,))
                cont = 0
                for row in cursor:
                    if cont % 2 == 0:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('par',))
                    else:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('impar',))
                    cont += 1

            elif opcao == 'Produto':
                db.cmd_reset_connection()
                tree_hist.delete(*tree_hist.get_children())
                cursor.execute("SELECT * FROM producao_reimp where prod_descricao like concat('%',%s,'%')",(pesq,))
                cont = 0
                for row in cursor:
                    if cont % 2 == 0:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('par',))
                    else:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('impar',))
                    cont += 1

            elif opcao == 'Usuário':
                db.cmd_reset_connection()
                tree_hist.delete(*tree_hist.get_children())
                cursor.execute("SELECT * FROM producao_reimp where usuario like concat('%',%s,'%')",(pesq,))
                cont = 0
                for row in cursor:
                    if cont % 2 == 0:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('par',))
                    else:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('impar',))
                    cont += 1

            elif opcao == 'Volume':
                db.cmd_reset_connection()
                tree_hist.delete(*tree_hist.get_children())
                cursor.execute("SELECT * FROM producao_reimp where lote like concat('%',%s,'%')",(pesq,))
                cont = 0
                for row in cursor:
                    if cont % 2 == 0:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('par',))
                    else:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('impar',))
                    cont += 1

        fr1 = Frame(frame4, bg='#ffffff')
        fr1.pack(side=TOP, fill=X)
        fr2 = Frame(frame4, bg='#222222')  # linha
        fr2.pack(side=TOP, fill=X, pady=5)
        fr3 = Frame(frame4, bg='#ffffff')
        fr3.pack(side=TOP, fill=X)
        fr4 = Frame(frame4, bg='#222222')  # linha
        fr4.pack(side=TOP, fill=X, pady=5)
        fr5 = Frame(frame4, bg='#ffffff')
        fr5.pack(side=TOP, fill=BOTH, expand=TRUE)    
        fr6 = Frame(frame4, bg='#ffffff')
        fr6.pack(side=TOP, fill=X)    

        #//////FRAME1
        image_invent = Image.open('img\\reimpressao2.png')
        resize_invent = image_invent.resize((60, 50))
        nova_image_invent = ImageTk.PhotoImage(resize_invent)
        lbl = Label(fr1, text=" Reimpressão de Etiquetas:", image=nova_image_invent, compound="left", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=0, column=1)
        fr1.grid_columnconfigure(0, weight=1)
        fr1.grid_columnconfigure(2, weight=1)
        #//////FRAME2 LINHA
        
        #//////FRAME3
        lbl_hist=Label(fr3, text='Histórico de Reimpressões: ', font=fonte_titulo_bold, bg='#ffffff', fg='#8b0000')
        lbl_hist.grid(row=0, column=1, padx=5, pady=(0,5))
      
        fr3.grid_columnconfigure(3, weight=1)
        
        clique_pesq = StringVar()
        lista_pesq = ['Corrida','Data','Produto','Usuário','Volume','Remover Filtro']

        opt_pesq = customtkinter.CTkOptionMenu(fr3, variable=clique_pesq, values=lista_pesq, width=160, **estilo_optionmenu,command=opt_pesq_imp)
        opt_pesq.grid(row=0, column=4, padx=2)
        opt_pesq.set('Pesquisar por...')

        ent_pesquisa = customtkinter.CTkEntry(fr3, **estilo_entry_padrao, width=160)
        ent_pesquisa.grid(row=0, column=5)
        ent_pesquisa.bind("<Return>", pesquisar_reimp_bind)

        image_pesq = Image.open('img\\pesq.png')
        resize_pesq = image_pesq.resize((30, 30))
        nova_image_pesq = ImageTk.PhotoImage(resize_pesq)
        btn_pesq = Button(fr3, image=nova_image_pesq, compound="top", bg='#FFFFFF', fg='#FFFFFF',
                        activebackground='#FFFFFF', borderwidth=0,
                        activeforeground="#FFFFFF", highlightthickness=4, relief=RIDGE, command=pesquisar_reimp,
                        font=fonte_botoes, cursor="hand2")
        btn_pesq.grid(row=0, column=6, padx=0)
        
        #//////FRAME4 LINHA

        #//////FRAME5
        style = ttk.Style()
        # style.theme_use('default')
        style.configure('Treeview',
                        background='#ffffff',
                        rowheight=24,
                        fieldbackground='#ffffff',
                        font=fonte_padrao)
        style.configure("Treeview.Heading",
                        foreground='#1d366c',
                        background="#ffffff",
                        font=fonte_padrao_bold)
        style.map('Treeview', background=[('selected', '#1D366C')])

        tree_hist = ttk.Treeview(fr5, selectmode='browse')
        vsb = ttk.Scrollbar(fr5, orient="vertical", command=tree_hist.yview)
        vsb.pack(side=RIGHT, fill='y')
        tree_hist.configure(yscrollcommand=vsb.set)
        vsbx = ttk.Scrollbar(fr5, orient="horizontal", command=tree_hist.xview)
        vsbx.pack(side=BOTTOM, fill='x')
        tree_hist.configure(xscrollcommand=vsbx.set)
        tree_hist.pack(side=LEFT, fill=BOTH, expand=True, anchor='n')
        tree_hist["columns"] = ("1", "2", "3", "4", "5", "6", "7")
        tree_hist['show'] = 'headings'
        tree_hist.column("#1", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#2", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#3", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#4", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#5", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#6", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#7", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.heading("1", text="Registro DB")
        tree_hist.heading("2", text="Produto")
        tree_hist.heading("3", text="Usuário")
        tree_hist.heading("4", text="Volume")
        tree_hist.heading("5", text="Corrida")
        tree_hist.heading("6", text="Peso")
        tree_hist.heading("7", text="Data")
        tree_hist.tag_configure('par', background='#e9e9e9')
        tree_hist.tag_configure('impar', background='#ffffff')
        #tree_hist.bind("<Double-1>", duploclique_tree_hist)
        fr5.grid_columnconfigure(0, weight=1)
        fr5.grid_columnconfigure(3, weight=1)

        #//////FRAME12
        bt_hist_imp = customtkinter.CTkButton(fr6, text='Histórico de Impressões', **estilo_botao_padrao_etiqueta, command=reimp)
        bt_hist_imp.grid(row=0, column=0, padx=10, pady=5)

        fr6.grid_columnconfigure(2, weight=1)

        atualizar_lista_reimpressao()
        mainloop()

    def hist_imp():
        def atualizar_lista_impressao():
            bt_select.configure(state='normal')
            db.cmd_reset_connection()
            tree_hist.delete(*tree_hist.get_children())
            cursor.execute("SELECT\
                A.id,\
                A.data,\
                A.hora,\
                A.usuario,\
                B.nome,\
                A.prod_descricao,\
                A.prod_codigo,\
                A.corrida,\
                A.lote,\
                A.peso,\
                IFNULL(A.envio_protheus,'') FROM producao A\
                inner join prod_tipo B on B.id = A.id_prod_tipo\
                where data = %s ORDER BY A.id DESC",(data,))
            cont = 0
            for row in cursor:
                if cont % 2 == 0:
                    tree_hist.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1], row[10]),
                                            tags=('par',))
                else:
                    tree_hist.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1], row[10]),
                                            tags=('impar',))
                cont += 1

        def setup_reimp():
            ent_data.insert(0, data)
            ent_data.configure(state='readonly')
            ent_prod_tipo.configure(state='readonly')
            ent_produto.configure(state='readonly')
            ent_cod.configure(state='readonly')
            ent_corrida.configure(state='readonly')
            ent_lote.configure(state='readonly')
            ent_seq.configure(state='readonly')
            ent_peso.configure(state='readonly')

        def impressao():
            cursor.execute('select * from op where numop = %s',(result2[10],))
            descricao = cursor.fetchone()
            global desc1
            desc1 = descricao[8].strip()
            global desc2        
            desc2 = descricao[9].strip()
            global desc3                
            desc3 = descricao[10].strip()


            data = ent_data.get()
            usuario = usuario_logado[1]
            prod_descricao = ent_produto.get()
            prod_codigo = ent_cod.get()
            corrida = ent_corrida.get()
            lote = ent_lote.get()
            peso = ent_peso.get()
            op = result2[10]

            if corrida == '' or peso == '':
                messagebox.showwarning('+Etiqueta:', 'Todos os campos devem ser preenchidos.', parent=root)
            else:
                cursor.execute("SELECT * FROM prod_tipo WHERE nome=%s", (ent_prod_tipo.get(),))
                produto = cursor.fetchone()[0]
                try:
                    cursor.execute("INSERT INTO producao_reimp (\
                        data,\
                        hora,\
                        usuario,\
                        id_prod_tipo,\
                        prod_descricao,\
                        prod_codigo,\
                        corrida,\
                        lote,\
                        peso,\
                        op)\
                        values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", (data, hora, usuario, produto, prod_descricao, prod_codigo, corrida, lote, peso, op))
                    db.commit()
                except Exception as e:
                    messagebox.showerror('+Etiqueta:', 'Erro ao gravar as informações na tabela (produção).', parent=root)
                    print(e)
                    return False

                def modelo1(): #Barra ou Endireitado
                    l = zpl.Label(114,75)
                    pos_y = 30
                    pos_x = 2
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"{desc1}",
                        char_height=4,
                        char_width=4,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"{desc2.replace('MM','mm')}",
                        char_height=4,
                        char_width=4,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"{desc3.replace('M','m')}",
                        char_height=4,
                        char_width=4,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 6
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Codigo {prod_codigo}",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Volume {lote}",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()


                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Corrida {corrida}",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Peso {peso} kg",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 5
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"{norma_abnt}",
                        char_height=2,
                        char_width=2,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 3
                    pos_x +=10
                    l.origin(pos_x,pos_y+5)
                    l.write_text(
                        "Registro",
                        char_height=2,
                        char_width=2,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    l.origin(pos_x-2,pos_y+7)
                    l.write_text(
                        f"{inmetro}",
                        char_height=3,
                        char_width=2,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    l.origin(42, 17)
                    l.write_barcode(height=60, barcode_type='C', check_digit='Y', orientation='R')
                    l.write_text(f'{prod_codigo}{lote}{corrida}{peso}')
                    l.endorigin()

                    pos_y += 8
                    l.origin(12, pos_y+10)
                    l.write_barcode(height=80, barcode_type='C', check_digit='Y', orientation='N')
                    l.write_text(f'{lote}')
                    l.endorigin()

                    
                    #print(height)

                    pg1 = l.dumpZPL()
                    #l.preview()

                    printer_name = win32print.GetDefaultPrinter ()
                    #
                    # raw_data could equally be raw PCL/PS read from
                    #  some print-to-file operation
                    #
                    if sys.version_info >= (3,):
                        #raw_data = bytes ("^XA^FO24,372^A0N,48,48^FB720,1,0,L,0^FDVERGALHAO CA-50^FS^FO24,420^A0N,48,48^FB720,1,0,L,0^FD8,00 MM ROLO^FS^FO24,468^A0N,48,48^FB720,1,0,L,0^FD//// CAUDA ////^FS^FO24,540^A0N,48,36^FB720,1,0,L,0^FDCodigo V1020B8.0^FS^FO24,588^A0N,48,36^FB720,1,0,L,0^FDVolume LHC14583^FS^FO24,636^A0N,48,36^FB720,1,0,L,0^FDCorrida 31244^FS^FO24,684^A0N,48,36^FB720,1,0,L,0^FDPeso 2,007 TON^FS^FO144,828^A0N,36,24^FB720,1,0,L,0^FDRegistro^FS^FO120,864^A0N,36,24^FB720,1,0,L,0^FD005713/2015^FS^FO504,204^BCB,60,N,N,Y,N^FDV1020B8.0 LHC1458331244 2,007^FS^FO144,996^BCN,80,S,N,Y,N^FDLHC14583^FS^XZ", "utf-8")
                        raw_data = bytes (pg1, "utf-8")
                        #print(raw_data) 
                    else:
                        raw_data = "This is a test"
                    hPrinter = win32print.OpenPrinter (printer_name)
                    try:
                        hJob = win32print.StartDocPrinter (hPrinter, 1, ("test of raw data", None, "RAW"))
                        try:
                            win32print.StartPagePrinter (hPrinter)
                            win32print.WritePrinter (hPrinter, raw_data)
                            win32print.EndPagePrinter (hPrinter)
                        finally:
                            win32print.EndDocPrinter (hPrinter)
                    finally:
                        win32print.ClosePrinter (hPrinter)

                def modelo2(): #Rolo
                    l = zpl.Label(114,75)
                    pos_y = 31
                    pos_x = 2
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"{desc1}",
                        char_height=4,
                        char_width=4,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"{desc2.replace('MM','mm')}",
                        char_height=4,
                        char_width=4,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        "//// CABECA ////",
                        char_height=4,
                        char_width=4,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 6
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Codigo {prod_codigo}",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Volume {lote}",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()


                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Corrida {corrida}",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Peso {peso} kg",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 5
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"{norma_abnt}",
                        char_height=2,
                        char_width=2,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 2
                    pos_x +=10
                    l.origin(pos_x,pos_y+5)
                    l.write_text(
                        "Registro",
                        char_height=2,
                        char_width=2,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    l.origin(pos_x-2,pos_y+7)
                    l.write_text(
                        f"{inmetro}",
                        char_height=3,
                        char_width=2,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()


                    l.origin(42, 17)
                    l.write_barcode(height=60, barcode_type='C', check_digit='Y', orientation='R')
                    l.write_text(f'{prod_codigo}{lote}{corrida}{peso}')
                    l.endorigin()

                    pos_y += 8
                    l.origin(12, pos_y+10)
                    l.write_barcode(height=80, barcode_type='C', check_digit='Y', orientation='N')
                    l.write_text(f'{lote}')
                    l.endorigin()

                    #print(height)

                    #print(l.dumpZPL())
                    pg1 = l.dumpZPL()
                    pg2 = pg1.replace('//// CABECA ////', '//// CAUDA ////')
                    uniao_paginas = pg1+pg2
                    #print(uniao_paginas)
                    
                    #l.preview()

                    printer_name = win32print.GetDefaultPrinter ()
                    #
                    # raw_data could equally be raw PCL/PS read from
                    #  some print-to-file operation
                    #
                    if sys.version_info >= (3,):
                        #raw_data = bytes ("^XA^FO24,372^A0N,48,48^FB720,1,0,L,0^FDVERGALHAO CA-50^FS^FO24,420^A0N,48,48^FB720,1,0,L,0^FD8,00 MM ROLO^FS^FO24,468^A0N,48,48^FB720,1,0,L,0^FD//// CAUDA ////^FS^FO24,540^A0N,48,36^FB720,1,0,L,0^FDCodigo V1020B8.0^FS^FO24,588^A0N,48,36^FB720,1,0,L,0^FDVolume LHC14583^FS^FO24,636^A0N,48,36^FB720,1,0,L,0^FDCorrida 31244^FS^FO24,684^A0N,48,36^FB720,1,0,L,0^FDPeso 2,007 TON^FS^FO144,828^A0N,36,24^FB720,1,0,L,0^FDRegistro^FS^FO120,864^A0N,36,24^FB720,1,0,L,0^FD005713/2015^FS^FO504,204^BCB,60,N,N,Y,N^FDV1020B8.0 LHC1458331244 2,007^FS^FO144,996^BCN,80,S,N,Y,N^FDLHC14583^FS^XZ", "utf-8")
                        raw_data = bytes (uniao_paginas, "utf-8")
                    else:
                        raw_data = "This is a test"
                    hPrinter = win32print.OpenPrinter (printer_name)
                    try:
                        hJob = win32print.StartDocPrinter (hPrinter, 1, ("test of raw data", None, "RAW"))
                        try:
                            win32print.StartPagePrinter (hPrinter)
                            win32print.WritePrinter (hPrinter, raw_data)
                            win32print.EndPagePrinter (hPrinter)
                        finally:
                            win32print.EndDocPrinter (hPrinter)
                    finally:
                        win32print.ClosePrinter (hPrinter)

                def modelo3(): #FM
                    l = zpl.Label(114,75)
                    pos_y = 30
                    pos_x = 2
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"{desc1.replace('MM','mm')}",
                        char_height=4,
                        char_width=4,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()
                    #6.50 MM|AÇO 1060
                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"{desc2}",
                        char_height=4,
                        char_width=4,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        "//// CABECA ////",
                        char_height=4,
                        char_width=4,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 6
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Codigo {prod_codigo}",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Volume {lote}",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()


                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Corrida {corrida}",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 4
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        f"Peso {peso} kg",
                        char_height=4,
                        char_width=3,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 5
                    l.origin(pos_x,pos_y)
                    l.write_text(
                        "",
                        char_height=2,
                        char_width=2,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    pos_y += 3
                    pos_x +=10
                    l.origin(pos_x,pos_y+5)
                    l.write_text(
                        "",
                        char_height=2,
                        char_width=2,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    l.origin(pos_x-2,pos_y+7)
                    l.write_text(
                        f"{inmetro}",
                        char_height=3,
                        char_width=2,
                        line_width=60,
                        justification='L',
                        )
                    l.endorigin()

                    l.origin(42, 17)
                    l.write_barcode(height=60, barcode_type='C', check_digit='Y', orientation='R')
                    l.write_text(f'{prod_codigo}{lote}{corrida}{peso}')
                    l.endorigin()

                    pos_y += 8
                    l.origin(12, pos_y+10)
                    l.write_barcode(height=80, barcode_type='C', check_digit='Y', orientation='N')
                    l.write_text(f'{lote}')
                    l.endorigin()

                    #l.preview()
                    #print(height)

                    pg1 = l.dumpZPL()
                    pg2 = pg1.replace('//// CABECA ////', '//// CAUDA ////')
                    uniao_paginas = pg1+pg2

                    printer_name = win32print.GetDefaultPrinter ()
                    #
                    # raw_data could equally be raw PCL/PS read from
                    #  some print-to-file operation
                    #
                    if sys.version_info >= (3,):
                        #raw_data = bytes ("^XA^FO24,372^A0N,48,48^FB720,1,0,L,0^FDVERGALHAO CA-50^FS^FO24,420^A0N,48,48^FB720,1,0,L,0^FD8,00 MM ROLO^FS^FO24,468^A0N,48,48^FB720,1,0,L,0^FD//// CAUDA ////^FS^FO24,540^A0N,48,36^FB720,1,0,L,0^FDCodigo V1020B8.0^FS^FO24,588^A0N,48,36^FB720,1,0,L,0^FDVolume LHC14583^FS^FO24,636^A0N,48,36^FB720,1,0,L,0^FDCorrida 31244^FS^FO24,684^A0N,48,36^FB720,1,0,L,0^FDPeso 2,007 TON^FS^FO144,828^A0N,36,24^FB720,1,0,L,0^FDRegistro^FS^FO120,864^A0N,36,24^FB720,1,0,L,0^FD005713/2015^FS^FO504,204^BCB,60,N,N,Y,N^FDV1020B8.0 LHC1458331244 2,007^FS^FO144,996^BCN,80,S,N,Y,N^FDLHC14583^FS^XZ", "utf-8")
                        raw_data = bytes (uniao_paginas, "utf-8")
                        #print(raw_data) 
                    else:
                        raw_data = "This is a test"
                    hPrinter = win32print.OpenPrinter (printer_name)
                    try:
                        hJob = win32print.StartDocPrinter (hPrinter, 1, ("test of raw data", None, "RAW"))
                        try:
                            win32print.StartPagePrinter (hPrinter)
                            win32print.WritePrinter (hPrinter, raw_data)
                            win32print.EndPagePrinter (hPrinter)
                        finally:
                            win32print.EndDocPrinter (hPrinter)
                    finally:
                        win32print.ClosePrinter (hPrinter)                    

                cursor.execute("SELECT inmetro from op WHERE numop = %s", (result2[10],))
                etq = cursor.fetchone()

                cursor.execute("select norma_abnt from norma_abnt")
                norma_abnt = cursor.fetchone()[0]

                if etq == None or etq == '':
                    inmetro = ''
                else:
                    inmetro = etq[0]
                if result2[4] == 'BARRA' or result2[4] == 'ENDIREITADO':
                    imp = threading.Thread(target=modelo1)
                    imp.start()
                elif result2[4] == 'ROLO':
                    imp = threading.Thread(target=modelo2)
                    imp.start()
                elif result2[4] == 'FIO MAQUINA':
                    imp = threading.Thread(target=modelo3)
                    imp.start()
                limpar_campos()

        def select():
            lista_select = tree_hist.focus()
            if lista_select == "":
                messagebox.showwarning('Reimpressão:', 'Selecione um item na lista!', parent=root)
            else:
                valor_lista = tree_hist.item(lista_select, "values")[0]
                try:
                    cursor.execute("SELECT\
                        A.id,\
                        A.data,\
                        A.hora,\
                        A.usuario,\
                        B.nome,\
                        A.prod_descricao,\
                        A.prod_codigo,\
                        A.corrida,\
                        A.lote,\
                        A.peso,\
                        A.op FROM producao A\
                        inner join prod_tipo B on B.id = A.id_prod_tipo\
                        WHERE A.id = %s",(valor_lista,))
                    global result2
                    result2 = cursor.fetchone()
                    print(result2)
                except Exception as e:
                    messagebox.showerror('Reimpressão:', 'Erro de conexão com o Banco de Dados.', parent=root)
                    print(e)
                    return False
            
            ent_data.configure(state='normal')
            ent_data.delete(0, END)
            ent_data.insert(0, data)
            ent_data.configure(state='readonly')
            
            ent_prod_tipo.configure(state='normal')
            ent_prod_tipo.delete(0, END)
            ent_prod_tipo.insert(0, result2[4])
            ent_prod_tipo.configure(state='readonly')        
            
            ent_produto.configure(state='normal')
            ent_produto.delete(0, END)
            ent_produto.insert(0, result2[5])
            ent_produto.configure(state='readonly')        

            ent_cod.configure(state='normal')
            ent_cod.delete(0, END)
            ent_cod.insert(0, result2[6])
            ent_cod.configure(state='readonly')        
                    
            ent_corrida.configure(state='normal')
            ent_corrida.delete(0, END)
            ent_corrida.insert(0, result2[7])
            ent_corrida.configure(state='readonly')        

            ent_lote.configure(state='normal')
            ent_lote.delete(0, END)
            ent_lote.insert(0, result2[8])
            ent_lote.configure(state='readonly')        
            
            ent_seq.configure(state='normal')
            ent_seq.delete(0, END)
            ent_seq.insert(0, result2[8])
            ent_seq.configure(state='readonly')        
        
            ent_peso.configure(state='normal')
            ent_peso.delete(0, END)
            ent_peso.insert(0, result2[9])
            ent_peso.configure(state='readonly')                        

        def limpar_campos():
            ent_data.configure(state='normal')
            ent_data.delete(0, END)
            ent_data.configure(state='readonly')
            
            ent_prod_tipo.configure(state='normal')
            ent_prod_tipo.delete(0, END)
            ent_prod_tipo.configure(state='readonly')        
            
            ent_produto.configure(state='normal')
            ent_produto.delete(0, END)
            ent_produto.configure(state='readonly')        

            ent_cod.configure(state='normal')
            ent_cod.delete(0, END)
            ent_cod.configure(state='readonly')        
                    
            ent_corrida.configure(state='normal')
            ent_corrida.delete(0, END)
            ent_corrida.configure(state='readonly')        

            ent_lote.configure(state='normal')
            ent_lote.delete(0, END)
            ent_lote.configure(state='readonly')        
            
            ent_seq.configure(state='normal')
            ent_seq.delete(0, END)
            ent_seq.configure(state='readonly')        
        
            ent_peso.configure(state='normal')
            ent_peso.delete(0, END)
            ent_peso.configure(state='readonly')                        

        def opt_pesq_imp(event):
            if clique_pesq.get() == 'Remover Filtro':
                ent_pesquisa.delete(0,END)
                opt_pesq.set('Pesquisar por...')
                atualizar_lista_impressao()
        
        def pesquisar_imp_bind(event):
            pesquisar_imp()

        def pesquisar_imp():
            opcao = clique_pesq.get()
            pesq = ent_pesquisa.get()
            if opcao == 'Corrida':
                db.cmd_reset_connection()
                tree_hist.delete(*tree_hist.get_children())
                cursor.execute("SELECT * FROM producao where corrida like concat('%',%s,'%')",(pesq,))
                cont = 0
                for row in cursor:
                    if cont % 2 == 0:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('par',))
                    else:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('impar',))
                    cont += 1

            elif opcao == 'Data':
                db.cmd_reset_connection()
                tree_hist.delete(*tree_hist.get_children())
                cursor.execute("SELECT * FROM producao where data like concat('%',%s,'%')",(pesq,))
                cont = 0
                for row in cursor:
                    if cont % 2 == 0:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('par',))
                    else:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('impar',))
                    cont += 1

            elif opcao == 'Produto':
                db.cmd_reset_connection()
                tree_hist.delete(*tree_hist.get_children())
                cursor.execute("SELECT * FROM producao where prod_descricao like concat('%',%s,'%')",(pesq,))
                cont = 0
                for row in cursor:
                    if cont % 2 == 0:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('par',))
                    else:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('impar',))
                    cont += 1

            elif opcao == 'Usuário':
                db.cmd_reset_connection()
                tree_hist.delete(*tree_hist.get_children())
                cursor.execute("SELECT * FROM producao where usuario like concat('%',%s,'%')",(pesq,))
                cont = 0
                for row in cursor:
                    if cont % 2 == 0:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('par',))
                    else:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('impar',))
                    cont += 1

            elif opcao == 'Volume':
                db.cmd_reset_connection()
                tree_hist.delete(*tree_hist.get_children())
                cursor.execute("SELECT * FROM producao where lote like concat('%',%s,'%')",(pesq,))
                cont = 0
                for row in cursor:
                    if cont % 2 == 0:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('par',))
                    else:
                        tree_hist.insert('', 'end', text=" ",
                                                values=(
                                                row[0], row[5], row[3], row[8], row[7], row[9], row[1]),
                                                tags=('impar',))
                    cont += 1

        fr1 = Frame(frame4, bg='#ffffff')
        fr1.pack(side=TOP, fill=X)
        fr2 = Frame(frame4, bg='#222222')  # linha
        fr2.pack(side=TOP, fill=X, pady=5)
        fr3 = Frame(frame4, bg='#ffffff')
        fr3.pack(side=TOP, fill=X)
        fr4 = Frame(frame4, bg='#222222')  # linha
        fr4.pack(side=TOP, fill=X, pady=5)
        fr5 = Frame(frame4, bg='#ffffff')
        fr5.pack(side=TOP, fill=X)
        fr6 = Frame(frame4, bg='#ffffff')
        fr6.pack(side=TOP, fill=X)    
        fr7 = Frame(frame4, bg='#222222')  # linha
        fr7.pack(side=TOP, fill=X, pady=5)
        fr8 = Frame(frame4, bg='#ffffff')
        fr8.pack(side=TOP, fill=X)    
        fr9 = Frame(frame4, bg='#222222')  # linha
        fr9.pack(side=TOP, fill=X, pady=5)
        fr10 = Frame(frame4, bg='#ffffff')
        fr10.pack(side=TOP, fill=X)    
        fr11 = Frame(frame4, bg='#ffffff')
        fr11.pack(side=TOP, fill=BOTH, expand=TRUE)    
        fr12 = Frame(frame4, bg='#ffffff')
        fr12.pack(side=TOP, fill=X)    

        #//////FRAME1
        image_invent = Image.open('img\\reimpressao2.png')
        resize_invent = image_invent.resize((60, 50))
        nova_image_invent = ImageTk.PhotoImage(resize_invent)
        lbl = Label(fr1, text=" Reimpressão de Etiquetas:", image=nova_image_invent, compound="left", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=0, column=1)
        fr1.grid_columnconfigure(0, weight=1)
        fr1.grid_columnconfigure(2, weight=1)
        #//////FRAME2 LINHA
        
        #//////FRAME3
        fr3.grid_columnconfigure(0, weight=1)
        fr3.grid_columnconfigure(6, weight=1)

        lbl = Label(fr3, text="Data:", font=fonte_padrao_bold, fg='#222222', bg='#ffffff')
        lbl.grid(row=0, column=1, sticky=W, padx=5)
        ent_data = customtkinter.CTkEntry(fr3, **estilo_entry_padrao, width=100)
        ent_data.grid(row=1, column=1, sticky=W, padx=5)

        lbl=Label(fr3, text='Linha de Produção: ', font=fonte_padrao_bold, bg='#ffffff', fg='#000000')
        lbl.grid(row=0, column=2, sticky="w", padx=5)
        ent_prod_tipo = customtkinter.CTkEntry(fr3, **estilo_entry_padrao, width=126)
        ent_prod_tipo.grid(row=1, column=2, sticky=W, padx=5)

        lbl=Label(fr3, text='Produtos: ', font=fonte_padrao_bold, bg='#ffffff', fg='#000000')
        lbl.grid(row=0, column=3, sticky="w", padx=5)
        ent_produto = customtkinter.CTkEntry(fr3, **estilo_entry_padrao, width=350)
        ent_produto.grid(row=1, column=3, sticky=W, padx=5)

        lbl = Label(fr3, text="Código:", font=fonte_padrao_bold, fg='#000000', bg='#ffffff')
        lbl.grid(row=0, column=4, sticky=W, padx=5)
        ent_cod = customtkinter.CTkEntry(fr3, **estilo_entry_padrao, width=180)
        ent_cod.grid(row=1, column=4, sticky=W, padx=5)


        #//////FRAME4 LINHA

        #//////FRAME5
        fr5.grid_columnconfigure(0, weight=1)
        fr5.grid_columnconfigure(6, weight=1)

        lbl = Label(fr5, text="Corrida:", font=fonte_padrao_bold, fg='#000000', bg='#ffffff')
        lbl.grid(row=0, column=2, sticky=W, padx=5)
        ent_corrida = customtkinter.CTkEntry(fr5, **estilo_entry_padrao_etiqueta,width=180)
        ent_corrida.grid(row=1, column=2, sticky=W, padx=5, pady=(0,5))
        
        lbl = Label(fr5, text="Lote|Volume:", font=fonte_padrao_bold, fg='#000000', bg='#ffffff')
        lbl.grid(row=0, column=3, sticky=W, padx=5)
        ent_lote = customtkinter.CTkEntry(fr5, **estilo_entry_padrao, width=180)
        ent_lote.grid(row=1, column=3, sticky=W, padx=5, pady=(0,5))

        lbl = Label(fr5, text="Sequência:", font=fonte_padrao_bold, fg='#000000', bg='#ffffff')
        lbl.grid(row=0, column=4, sticky=W, padx=5)
        ent_seq = customtkinter.CTkEntry(fr5, **estilo_entry_padrao_etiqueta, width=180)
        ent_seq.grid(row=1, column=4, sticky=W, padx=5, pady=(0,5))

        lbl = Label(fr5, text="Peso (KG):", font=fonte_padrao_bold, fg='#000000', bg='#ffffff')
        lbl.grid(row=0, column=5, sticky=W, padx=5)
        ent_peso = customtkinter.CTkEntry(fr5, **estilo_entry_padrao_etiqueta, width=200)
        ent_peso.grid(row=1, column=5, sticky=W, padx=5)

        #//////FRAME6
        fr6.grid_columnconfigure(0, weight=1)
        fr6.grid_columnconfigure(6, weight=1)


        #//////FRAME7 LINHA
        
        #//////FRAME8
        fr8.grid_columnconfigure(0, weight=1)
        fr8.grid_columnconfigure(6, weight=1)
        global bt_imp
        bt_imp = customtkinter.CTkButton(fr8, text='Imprimir (F1)', **estilo_botao_padrao_etiqueta, command=impressao)
        bt_imp.grid(row=1, column=1)
        root.bind('<F1>', lambda event: impressao())
        #root.bind('<F2>', lambda event: capturar_peso())
        
        #//////FRAME9 LINHA
        
        #//////FRAME10
        lbl_hist=Label(fr10, text='Histórico de Impressões: ', font=fonte_titulo_bold, bg='#ffffff', fg='#222222')
        lbl_hist.grid(row=0, column=1, padx=5, pady=(0,5))

        bt_select = customtkinter.CTkButton(fr10, text='Selecionar Item', **estilo_botao_padrao_etiqueta, command=select)
        bt_select.grid(row=0, column=2)
        
        fr10.grid_columnconfigure(3, weight=1)
        
        clique_pesq = StringVar()
        lista_pesq = ['Corrida','Data','Produto','Usuário','Volume','Remover Filtro']

        opt_pesq = customtkinter.CTkOptionMenu(fr10, variable=clique_pesq, values=lista_pesq, width=160, **estilo_optionmenu,command=opt_pesq_imp)
        opt_pesq.grid(row=0, column=4, padx=2)
        opt_pesq.set('Pesquisar por...')
        
        ent_pesquisa = customtkinter.CTkEntry(fr10, **estilo_entry_padrao, width=160)
        ent_pesquisa.grid(row=0, column=5)
        ent_pesquisa.bind("<Return>", pesquisar_imp_bind)

        image_pesq = Image.open('img\\pesq.png')
        resize_pesq = image_pesq.resize((30, 30))
        nova_image_pesq = ImageTk.PhotoImage(resize_pesq)
        btn_pesq = Button(fr10, image=nova_image_pesq, compound="top", bg='#FFFFFF', fg='#FFFFFF',
                        activebackground='#FFFFFF', borderwidth=0,
                        activeforeground="#FFFFFF", highlightthickness=4, relief=RIDGE, command=pesquisar_imp,
                        font=fonte_botoes, cursor="hand2")
        btn_pesq.grid(row=0, column=6, padx=0)

        #//////FRAME11
        style = ttk.Style()
        # style.theme_use('default')
        style.configure('Treeview',
                        background='#ffffff',
                        rowheight=24,
                        fieldbackground='#ffffff',
                        font=fonte_padrao)
        style.configure("Treeview.Heading",
                        foreground='#1d366c',
                        background="#ffffff",
                        font=fonte_padrao_bold)
        style.map('Treeview', background=[('selected', '#1D366C')])

        tree_hist = ttk.Treeview(fr11, selectmode='browse')
        vsb = ttk.Scrollbar(fr11, orient="vertical", command=tree_hist.yview)
        vsb.pack(side=RIGHT, fill='y')
        tree_hist.configure(yscrollcommand=vsb.set)
        vsbx = ttk.Scrollbar(fr11, orient="horizontal", command=tree_hist.xview)
        vsbx.pack(side=BOTTOM, fill='x')
        tree_hist.configure(xscrollcommand=vsbx.set)
        tree_hist.pack(side=LEFT, fill=BOTH, expand=True, anchor='n')
        tree_hist["columns"] = ("1", "2", "3", "4", "5", "6", "7", "8")
        tree_hist['show'] = 'headings'
        tree_hist.column("#1", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#2", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#3", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#4", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#5", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#6", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#7", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.column("#8", anchor='c', minwidth=50, width=100, stretch = True)
        tree_hist.heading("1", text="Registro DB")
        tree_hist.heading("2", text="Produto")
        tree_hist.heading("3", text="Usuário")
        tree_hist.heading("4", text="Volume")
        tree_hist.heading("5", text="Corrida")
        tree_hist.heading("6", text="Peso")
        tree_hist.heading("7", text="Data")
        tree_hist.heading("8", text="Envio Protheus")
        tree_hist.tag_configure('par', background='#e9e9e9')
        tree_hist.tag_configure('impar', background='#ffffff')
        #tree_hist.bind("<Double-1>", duploclique_tree_hist)
        fr11.grid_columnconfigure(0, weight=1)
        fr11.grid_columnconfigure(3, weight=1)

        #//////FRAME12
      
        bt_hist_reimp = customtkinter.CTkButton(fr12, text='Histórico de Reimpressões', **estilo_botao_padrao_vermelho, command=hist_reimp)
        bt_hist_reimp.grid(row=0, column=1, padx=10, pady=5)
        
        fr12.grid_columnconfigure(2, weight=1)

        atualizar_lista_impressao()
        setup_reimp()
        mainloop()

    for widget in frame4.winfo_children():
        widget.destroy()

    hist_imp()

def cadastro_usuarios():
    global controle_loop
    controle_loop = ''
        
    for widget in frame4.winfo_children():
        widget.destroy()

    def atualizar_lista_usuarios():
        db.cmd_reset_connection()
        tree_cadastro.delete(*tree_cadastro.get_children())
        cursor.execute("SELECT * from usuarios ORDER BY nome")
        cont = 0
        for row in cursor:
            if row[4] == 1:
                acesso = 'Administrador'
            else:
                acesso = 'Operador'
            if cont % 2 == 0:
                tree_cadastro.insert('', 'end', text=" ",
                                        values=(
                                        row[0], row[1], row[2], acesso),
                                        tags=('par',))
            else:
                tree_cadastro.insert('', 'end', text=" ",
                                        values=(
                                        row[0], row[1], row[2], acesso),
                                        tags=('impar',))
            cont += 1

    def salvar():
        nome = ent_nome.get().upper()
        usuario = ent_user.get()
        senha = ent_senha.get()
        gestor = chk_var1.get()
        if nome == '' or usuario == '' or senha == ''or gestor == '':
            messagebox.showwarning('Cadastro de Usuários:', 'Todos os campos devem ser preenchidos.', parent=root)
        else:
            x=''
            try:
                autent = usuario_logado[2]+':'+usuario_logado[3]
                autent_b = autent.encode()
                autorizacao = base64.b64encode(autent_b)
                #url = 'http://192.168.1.18:8689/rest/api/oauth2/v1/token' //Desenv
                url = 'http://192.168.1.18:8683/rest/api/oauth2/v1/token'
                headers = {
                    'POST': '/rest/api/oauth2/v1/token',
                    'Host': 'http://192.168.1.18:8683',
                    'Accept': 'application/json',
                    'Authentication': 'BASIC '+ str(autorizacao),
                }
                parametros = {
                    'grant_type':'password',
                    'username':f'{usuario}',
                    'password':f'{senha}',
                }
                x = requests.post(url, headers=headers, params=parametros, timeout=2)
                x = x.status_code
                print(x)
            except Exception as e:
                messagebox.showerror('Cadastro de Usuários:', 'Erro de conexão com o Protheus.', parent=root)
                print(e)
                return False
            if x == 200 or x == 201:
                cursor.execute("SELECT * FROM usuarios WHERE usuario=%s", (usuario,))
                verifica_usuario = cursor.fetchone()
               
                if verifica_usuario == None:
                    try:
                        cursor.execute(
                            "INSERT INTO usuarios (nome, usuario, senha, gestor) values(%s,%s,%s,%s)", (nome,usuario,senha,gestor))
                        db.commit()
                    except Exception as e:
                        messagebox.showerror('Cadastro de Usuários:', 'Erro de conexão com o Banco de Dados.', parent=root)
                        print(e)
                        return False

                    messagebox.showinfo('Cadastro de Usuários:', 'Cadastro efetuado com sucesso.', parent=root)
                    cadastro_usuarios()
                else:
                    messagebox.showerror('Cadastro de Usuários', 'Usuário já cadastrado:', parent=root)
            else:
                messagebox.showwarning('Cadastro de Usuários', 'Utilize o login e senha do Protheus.', parent=root)

    def editar():
        def cancelar():
            cadastro_usuarios()
        
        def confirmar():
            nome = ent_nome.get().upper()
            usuario = ent_user.get()
            senha = ent_senha.get()
            gestor = chk_var1.get()

            if nome == '' or usuario == '' or senha == ''or gestor == '':
                messagebox.showwarning('Cadastro de Usuários:', 'Todos os campos devem ser preenchidos.', parent=root)
            else:
                x=''
                try:
                    url = 'http://192.168.1.18:8683/rest/api/oauth2/v1/token'
                    parametros = {
                        'grant_type':'password',
                        'username':f'{usuario}',
                        'password':f'{senha}',
                    }
                    x = requests.post(url, params=parametros, timeout=2)
                    x = x.status_code
                except Exception as e:
                    messagebox.showerror('Cadastro de Usuários:', 'Erro de conexão com o Protheus.', parent=root)
                    print(e)
                    return False
                if x == 200:
                    try:
                        cursor.execute("UPDATE usuarios SET nome = %s, usuario= %s, senha= %s, gestor= %s WHERE id = %s", (nome,usuario,senha,gestor,id))
                        db.commit()
                    except Exception as e:
                        messagebox.showerror('Cadastro de Usuários:', 'Erro de conexão com o Banco de Dados.', parent=root)
                        print(e)
                        return False
                    messagebox.showinfo('Cadastro de Usuários:', 'Usuário editado com sucesso.', parent=root)
                    cadastro_usuarios()
                else:
                    messagebox.showwarning('Cadastro de Usuários', 'Utilize o login e senha do Protheus.', parent=root)


        lista_select = tree_cadastro.focus()
        valor_lista = tree_cadastro.item(lista_select, "values")
        if valor_lista == '':
            print('Mensagem de Erro')
        else:
            btn_salvar_produto.grid_remove()
            btn_editar_produto.grid_remove()

            image_confirmar = Image.open('img\\salvar.png')
            resize_confirmar = image_confirmar.resize((28, 32))
            nova_image_confirmar = ImageTk.PhotoImage(resize_confirmar)
            btn_confirmar_produto =Button(fr8, image=nova_image_confirmar, text="Confirmar", compound="top",
                                font=fonte_padrao_bold, bg='#ffffff', fg='#222222', command=confirmar,
                                borderwidth=0, relief=RIDGE, activebackground="#ffffff", activeforeground="#1d366c", cursor="hand2")
            btn_confirmar_produto.photo = nova_image_confirmar
            btn_confirmar_produto.grid(row=0, column=1, padx=4)

            image_cancelar = Image.open('img\\cancelar.png')
            resize_cancelar = image_cancelar.resize((28, 32))
            nova_image_cancelar = ImageTk.PhotoImage(resize_cancelar)
            btn_cancelar_produto =Button(fr8, image=nova_image_cancelar, text="Cancelar", compound="top",
                                font=fonte_padrao_bold, bg='#ffffff', fg='#222222', command=cancelar,
                                borderwidth=0, relief=RIDGE, activebackground="#ffffff", activeforeground="#1d366c", cursor="hand2")
            btn_cancelar_produto.photo = nova_image_cancelar
            btn_cancelar_produto.grid(row=0, column=2, padx=4)

            id= valor_lista[0]
            ent_nome.delete(0,END)
            ent_nome.insert(0, valor_lista[1])
            ent_user.delete(0,END)
            ent_user.insert(0, valor_lista[2])
            if valor_lista[3] == 'Administrador':
                chk_var1.set(1)
            else:
                chk_var1.set(0)


    fr1 = Frame(frame4, bg='#ffffff')
    fr1.pack(side=TOP, fill=X)
    fr2 = Frame(frame4, bg='#222222')  # linha
    fr2.pack(side=TOP, fill=X)
    fr3 = Frame(frame4, bg='#ffffff')
    fr3.pack(side=TOP, fill=X)
    fr4 = Frame(frame4, bg='#ffffff')
    fr4.pack(side=TOP, fill=X)
    fr5 = Frame(frame4, bg='#222222')  # linha
    fr5.pack(side=TOP, fill=X)
    fr6 = Frame(frame4, bg='#ffffff')
    fr6.pack(side=TOP, fill=BOTH, expand=TRUE)
    fr7 = Frame(frame4, bg='#222222')  # linha
    fr7.pack(side=TOP, fill=X)
    fr8 = Frame(frame4, bg='#ffffff')
    fr8.pack(side=TOP, fill=X)   

    #//////FRAME1
    image_invent = Image.open('img\\cadastro2.png')
    resize_invent = image_invent.resize((50, 50))
    nova_image_invent = ImageTk.PhotoImage(resize_invent)
    lbl = Label(fr1, text=" Cadastro | Usuários", image=nova_image_invent, compound="left", font=fonte_titulo_bold, fg='#222222', bg='#ffffff')
    lbl.grid(row=0, column=1)
    fr1.grid_columnconfigure(0, weight=1)
    fr1.grid_columnconfigure(2, weight=1)
    #//////FRAME2 LINHA
    
    #//////FRAME3
    lbl_nome=Label(fr3, text='Nome:', font=fonte_padrao, bg='#ffffff', fg='#000000')
    lbl_nome.grid(row=0, column=1, sticky="w")
    ent_nome = customtkinter.CTkEntry(fr3, **estilo_entry_padrao, width=340)
    ent_nome.grid(row=1, column=1, pady=3)        
    ent_nome.focus_force()
    
    lbl_user=Label(fr3, text='Usuário:', font=fonte_padrao, bg='#ffffff', fg='#000000')
    lbl_user.grid(row=2, column=1, sticky="w")
    ent_user = customtkinter.CTkEntry(fr3, **estilo_entry_padrao, width=340)
    ent_user.grid(row=3, column=1, pady=3)        

    lbl_senha=Label(fr3, text='Senha:', font=fonte_padrao, bg='#ffffff', fg='#000000')
    lbl_senha.grid(row=4, column=1, sticky="w")
    ent_senha = customtkinter.CTkEntry(fr3, **estilo_entry_padrao, width=340, show='*')
    ent_senha.grid(row=5, column=1, pady=3)        
    
    fr3.grid_columnconfigure(0, weight=1)
    fr3.grid_columnconfigure(6, weight=1)    

    chk_var1 = IntVar()
    chk_1 = customtkinter.CTkCheckBox(fr4, text='Gestor\Líder', variable=chk_var1, onvalue=1, offvalue=0, font=fonte_padrao, bg_color='#ffffff', fg_color='#2a2d2e', hover_color='#2a2d2e', text_color='#2a2d2e', border_color='#2a2d2e')
    chk_1.grid(row=4, column=2, padx=10, pady=10)
    
    fr4.grid_columnconfigure(0, weight=1)
    fr4.grid_columnconfigure(3, weight=1)
    
    #//FRAME5 LINHA

    #//FRAME6
    style = ttk.Style()
    # style.theme_use('default')
    style.configure('Treeview',
                    background='#ffffff',
                    rowheight=24,
                    fieldbackground='#ffffff',
                    font=fonte_padrao)
    style.configure("Treeview.Heading",
                    foreground='#000000',
                    background="#ffffff",
                    font=fonte_padrao)
    style.map('Treeview', background=[('selected', '#1d366c')])

    tree_cadastro = ttk.Treeview(fr6, selectmode='browse')
    vsb = ttk.Scrollbar(fr6, orient="vertical", command=tree_cadastro.yview)
    vsb.pack(side=RIGHT, fill='y')
    #tree_cadastro.configure(yscrollcommand=vsb.set)
    vsbx = ttk.Scrollbar(fr6, orient="horizontal", command=tree_cadastro.xview)
    vsbx.pack(side=BOTTOM, fill='x')
    tree_cadastro.configure(xscrollcommand=vsbx.set)
    tree_cadastro.pack(side=LEFT, fill=BOTH, expand=True, anchor='n')
    tree_cadastro["columns"] = ("1", "2", "3", "4")
    tree_cadastro['show'] = 'headings'
    tree_cadastro.column("#1", anchor='c', minwidth=50, width=100, stretch = True)    
    tree_cadastro.column("#2", anchor='c', minwidth=50, width=100, stretch = True)    
    tree_cadastro.column("#3", anchor='c', minwidth=50, width=100, stretch = True)    
    tree_cadastro.column("#4", anchor='c', minwidth=50, width=100, stretch = True)    
    tree_cadastro.heading("1", text="ID")
    tree_cadastro.heading("2", text="Nome")
    tree_cadastro.heading("3", text="Usuário")
    tree_cadastro.heading("4", text="Nível|Acesso")
    tree_cadastro.tag_configure('par', background='#e9e9e9')
    tree_cadastro.tag_configure('impar', background='#ffffff')
    
    #//FRAME7 LINHA

    #//FRAME8
    image_salvar = Image.open('img\\salvar.png')
    resize_salvar = image_salvar.resize((28, 32))
    nova_image_salvar = ImageTk.PhotoImage(resize_salvar)
    btn_salvar_produto =Button(fr8, image=nova_image_salvar, text="Salvar", compound="top",
                        font=fonte_padrao_bold, bg='#ffffff', fg='#222222', command=salvar,
                        borderwidth=0, relief=RIDGE, activebackground="#ffffff", activeforeground="#1d366c", cursor="hand2")
    btn_salvar_produto.photo = nova_image_salvar
    btn_salvar_produto.grid(row=0, column=1, padx=4)

    image_editar = Image.open('img\\editar.png')
    resize_editar = image_editar.resize((28, 32))
    nova_image_editar = ImageTk.PhotoImage(resize_editar)
    btn_editar_produto =Button(fr8, image=nova_image_editar, text="Editar", compound="top",
                        font=fonte_padrao_bold, bg='#ffffff', fg='#222222', command=editar,
                        borderwidth=0, relief=RIDGE, activebackground="#ffffff", activeforeground="#1d366c", cursor="hand2")
    btn_editar_produto.photo = nova_image_editar
    btn_editar_produto.grid(row=0, column=2, padx=4)


    fr8.grid_columnconfigure(0, weight=1)
    fr8.grid_columnconfigure(3, weight=1)

    atualizar_lista_usuarios()
    mainloop()

def home():
    global controle_loop
    controle_loop = 'home'

    for widget in frame4.winfo_children():
        widget.destroy()

    def duplo_clique_tree_principal(event): #\\ Ao dar um duplo clique na home e caso o status for Erro e o usuário for um gestor, o sistema vai reenviar o item para validação.

        lista_select = tree_principal.focus()
        valor_lista = tree_principal.item(lista_select, "values")
        status = valor_lista[7]
        id = valor_lista[0]
        
        if status == 'Erro' and usuario_logado[4] == 1:
            try:
                cursor.execute("Select * from producao where id = %s",(id,))
                resultado = cursor.fetchone()
                lista = []
                dic = {}
                peso = float(resultado[9])
                resto = peso%1000
                if resto == 0:
                    peso =int(str((peso/1000)).rstrip('.0'))
                else:
                    peso = peso/1000
                ano =(resultado[1])[6:]
                mes =(resultado[1])[3:5]
                dia =(resultado[1])[0:2]                    
                data_b = ano+mes+dia
                dic['op'] = resultado[10]
                dic['corrida'] = resultado[7]
                dic['volume'] = resultado[8]
                dic['data'] = data_b
                dic['operador'] = resultado[3]
                dic['peso'] = peso
                lista.append(dic)
                retorno = ''

                url = 'http://192.168.1.18:8683/rest/SIMEC_PCP/producao'
                r = requests.post(url, json = lista, timeout=0.9)
                retorno = json.loads(r.content)
            except Exception as e:
                lbl_status_erro1.configure(text= f'Erro - Home(Função: duplo clique) / {e}')
                print('Erro - Home(Função: duplo clique) /', e)
                pass
            for i in retorno:
                if i['status'] == True:
                    cursor.execute("update producao set envio_protheus = %s, envio_data = %s, envio_hora = %s, envio_tentativa = %s WHERE lote = %s",('Processando',data,hora,1,i['volume'],))
                    messagebox.showinfo('Forçar reenvio:', 'Registro enviado com sucesso!', parent=root)
                db.commit()
                atualizar_lista_principal()

    def opt_pesq_imp(event):
        if clique_pesq.get() == 'Remover Filtro':
            ent_pesquisa.delete(0,END)
            opt_pesq.set('Pesquisar por...')
            atualizar_lista_principal()
        elif clique_pesq.get() == 'Erro de Envio':
            db.cmd_reset_connection()
            tree_principal.delete(*tree_principal.get_children())
            cursor.execute("SELECT * FROM producao where envio_protheus = 'Erro' order by ID DESC")
            cont = 0
            for row in cursor:
                if cont % 2 == 0:
                    tree_principal.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1], row[11]),
                                            tags=('par',))
                else:
                    tree_principal.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1], row[11]),
                                            tags=('impar',))
                cont += 1

    def pesquisar_imp_bind(event):
        pesquisar_imp()

    def pesquisar_imp():
        opcao = clique_pesq.get()
        pesq = ent_pesquisa.get()
        if opcao == 'Corrida':
            db.cmd_reset_connection()
            tree_principal.delete(*tree_principal.get_children())
            cursor.execute("SELECT * FROM producao where corrida like concat('%',%s,'%')",(pesq,))
            cont = 0
            for row in cursor:
                if cont % 2 == 0:
                    tree_principal.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1], row[11]),
                                            tags=('par',))
                else:
                    tree_principal.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1], row[11]),
                                            tags=('impar',))
                cont += 1

        elif opcao == 'Data':
            db.cmd_reset_connection()
            tree_principal.delete(*tree_principal.get_children())
            cursor.execute("SELECT * FROM producao where data like concat('%',%s,'%')",(pesq,))
            cont = 0
            for row in cursor:
                if cont % 2 == 0:
                    tree_principal.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1], row[11]),
                                            tags=('par',))
                else:
                    tree_principal.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1], row[11]),
                                            tags=('impar',))
                cont += 1

        elif opcao == 'Produto':
            db.cmd_reset_connection()
            tree_principal.delete(*tree_principal.get_children())
            cursor.execute("SELECT * FROM producao where prod_descricao like concat('%',%s,'%')",(pesq,))
            cont = 0
            for row in cursor:
                if cont % 2 == 0:
                    tree_principal.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1], row[11]),
                                            tags=('par',))
                else:
                    tree_principal.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1], row[11]),
                                            tags=('impar',))
                cont += 1

        elif opcao == 'Usuário':
            db.cmd_reset_connection()
            tree_principal.delete(*tree_principal.get_children())
            cursor.execute("SELECT * FROM producao where usuario like concat('%',%s,'%')",(pesq,))
            cont = 0
            for row in cursor:
                if cont % 2 == 0:
                    tree_principal.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1], row[11]),
                                            tags=('par',))
                else:
                    tree_principal.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1], row[11]),
                                            tags=('impar',))
                cont += 1

        elif opcao == 'Volume':
            db.cmd_reset_connection()
            tree_principal.delete(*tree_principal.get_children())
            cursor.execute("SELECT * FROM producao where lote like concat('%',%s,'%')",(pesq,))
            cont = 0
            for row in cursor:
                if cont % 2 == 0:
                    tree_principal.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1], row[11]),
                                            tags=('par',))
                else:
                    tree_principal.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1], row[11]),
                                            tags=('impar',))
                cont += 1

    def loop_home():
        if controle_loop == 'home': #// Loop para ficar atualizando a pagina home
            root.after(30000, atualizar_lista_principal)

    def atualizar_lista_principal():
        if controle_loop == 'home':
            #print('Home - atualizar_lista_principal')
            db.cmd_reset_connection()
            tree_principal.delete(*tree_principal.get_children())
            cursor.execute("SELECT\
                A.id,\
                A.data,\
                A.hora,\
                A.usuario,\
                B.nome,\
                A.prod_descricao,\
                A.prod_codigo,\
                A.corrida,\
                A.lote,\
                A.peso,\
                IFNULL(A.envio_protheus,'') FROM producao A\
                inner join prod_tipo B on B.id = A.id_prod_tipo\
                where data = %s ORDER BY A.id DESC",(data,))
            cont = 0
            for row in cursor:
                if cont % 2 == 0:
                    tree_principal.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1], row[10]),
                                            tags=('par',))
                else:
                    tree_principal.insert('', 'end', text=" ",
                                            values=(
                                            row[0], row[5], row[3], row[8], row[7], row[9], row[1], row[10]),
                                            tags=('impar',))
                cont += 1
            loop_home()
    
    fr0 = customtkinter.CTkFrame(frame4, corner_radius=10, fg_color='#ffffff', border_width=2, border_color='#2a2d2e')
    fr0.pack(padx=4, pady=10, fill="both", expand=True)

    fr1 = Frame(fr0, bg='#ffffff')
    fr1.pack(padx=10, pady=5, fill="x", expand=False, side=TOP)
    fr2 = Frame(fr0, bg='#2a2d2e') #/// LINHA
    fr2.pack(padx=10, pady=0, fill="x", expand=False, side=TOP)
    fr3 = Frame(fr0, bg='#ffffff')
    fr3.pack(padx=10, pady=5, fill=BOTH, expand=TRUE, side=TOP)    
    
    lbl_logado = Label(fr1, text="Histórico de Impressões", bg='#ffffff', fg="#222222", font=fonte_titulo_bold)
    lbl_logado.grid(row=0, column=0, padx=2)

    fr1.grid_columnconfigure(1, weight=1)

    clique_pesq = StringVar()
    lista_pesq = ['Corrida','Data','Erro de Envio','Produto','Usuário','Volume','Remover Filtro']

    opt_pesq = customtkinter.CTkOptionMenu(fr1, variable=clique_pesq, values=lista_pesq, width=160, **estilo_optionmenu,command=opt_pesq_imp)
    opt_pesq.grid(row=0, column=2, padx=2)
    opt_pesq.set('Pesquisar por...')
    
    ent_pesquisa = customtkinter.CTkEntry(fr1, **estilo_entry_padrao, width=160)
    ent_pesquisa.grid(row=0, column=3)
    ent_pesquisa.bind("<Return>", pesquisar_imp_bind)

    image_pesq = Image.open('img\\pesq.png')
    resize_pesq = image_pesq.resize((30, 30))
    nova_image_pesq = ImageTk.PhotoImage(resize_pesq)
    btn_pesq = Button(fr1, image=nova_image_pesq, compound="top", bg='#FFFFFF', fg='#FFFFFF',
                    activebackground='#FFFFFF', borderwidth=0,
                    activeforeground="#FFFFFF", highlightthickness=4, relief=RIDGE, command=pesquisar_imp, textvariable=pesquisar_imp_bind,
                    font=fonte_botoes, cursor="hand2")
    btn_pesq.grid(row=0, column=4, padx=0)


    style = ttk.Style()
    # style.theme_use('default')
    style.configure('Treeview',
                    background='#ffffff',
                    rowheight=24,
                    fieldbackground='#ffffff',
                    font=fonte_padrao)
    style.configure("Treeview.Heading",
                    foreground='#1d366c',
                    background="#ffffff",
                    font=fonte_padrao_bold)
    style.map('Treeview', background=[('selected', '#1D366C')])
    global tree_principal
    tree_principal = ttk.Treeview(fr3, selectmode='browse')
    vsb = ttk.Scrollbar(fr3, orient="vertical", command=tree_principal.yview)
    vsb.pack(side=RIGHT, fill='y')
    tree_principal.configure(yscrollcommand=vsb.set)
    vsbx = ttk.Scrollbar(fr3, orient="horizontal", command=tree_principal.xview)
    vsbx.pack(side=BOTTOM, fill='x')
    tree_principal.configure(xscrollcommand=vsbx.set)
    tree_principal.pack(side=LEFT, fill=BOTH, expand=True, anchor='n')
    tree_principal["columns"] = ("1", "2", "3", "4", "5", "6", "7", "8")
    tree_principal['show'] = 'headings'
    tree_principal.column("#1", anchor='c', minwidth=50, width=100, stretch = True)
    tree_principal.column("#2", anchor='c', minwidth=50, width=100, stretch = True)
    tree_principal.column("#3", anchor='c', minwidth=50, width=100, stretch = True)
    tree_principal.column("#4", anchor='c', minwidth=50, width=100, stretch = True)
    tree_principal.column("#5", anchor='c', minwidth=50, width=100, stretch = True)
    tree_principal.column("#6", anchor='c', minwidth=50, width=100, stretch = True)
    tree_principal.column("#7", anchor='c', minwidth=50, width=100, stretch = True)
    tree_principal.column("#8", anchor='c', minwidth=50, width=100, stretch = True)
    tree_principal.heading("1", text="Registro DB")
    tree_principal.heading("2", text="Produto")
    tree_principal.heading("3", text="Usuário")
    tree_principal.heading("4", text="Volume")
    tree_principal.heading("5", text="Corrida")
    tree_principal.heading("6", text="Peso")
    tree_principal.heading("7", text="Data")
    tree_principal.heading("8", text="Envio Protheus")    
    tree_principal.tag_configure('par', background='#e9e9e9')
    tree_principal.tag_configure('impar', background='#ffffff')
    tree_principal.bind("<Double-1>", duplo_clique_tree_principal)
    fr3.grid_columnconfigure(0, weight=1)
    fr3.grid_columnconfigure(3, weight=1)

    atualizar_lista_principal()
    root.mainloop()

def verifica_senha_protheus():
    x=''
    autent = usuario_logado[2]+':'+usuario_logado[3]
    autent_b = autent.encode()
    autorizacao = base64.b64encode(autent_b)

    try:
        #url = 'http://192.168.1.18:8689/rest/api/oauth2/v1/token' //Desenv
        url = 'http://192.168.1.18:8683/rest/api/oauth2/v1/token'
        headers = {
            'POST': '/rest/api/oauth2/v1/token',
            'Host': 'http://192.168.1.18:8683',
            'Accept': 'application/json',
            'Authentication': 'BASIC '+ str(autorizacao),
        }
        parametros = {
            'grant_type':'password',
            'username':f'{usuario_logado[2]}',
            'password':f'{usuario_logado[3]}',
        }
        x = requests.post(url, headers=headers, params=parametros, timeout=2)
        x = x.status_code
    except Exception as e:
        lbl_status_erro1.configure(text= f'Erro - Login(Função: verifica_senha_protheus) / {e}')
        print('Erro - Login(Função: verifica_senha_protheus) /', e)
        pass

    if x == 401:
        atualiza_senha_protheus()

def bloqueio_botao(): #\\Bloqueia o botao de imprimir por um tempo.
    btn_etiqueta.configure(state='disabled')
    time.sleep(4)
    btn_etiqueta.configure(state='normal')

def login():
    global controle_loop
    controle_loop = ''
        
    root2 = Toplevel(root)
    root2.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
    root2.unbind_class("Button", "<Key-space>")
    root2.focus_force()
    root2.grab_set()

    window_width = 400
    window_height = 300
    screen_width = root2.winfo_screenwidth()
    screen_height = root2.winfo_screenheight() - 70
    x_cordinate = int((screen_width / 2) - (window_width / 2))
    y_cordinate = int((screen_height / 2) - (window_height / 2))
    root2.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
    root2.resizable(0, 0)
    root2.configure(bg='#ffffff')
    root2.title(titulo)
    root2.overrideredirect(True)
    root2.iconbitmap('img\\icone.ico')

    
    #///////////////////////// FUNÇÕES
    def logar():
        usuario = ent_usuario.get()
        senha = ent_senha.get()

        if usuario == "" or senha == "":
            messagebox.showwarning('Login:', 'Digite o Usuário ou Senha.', parent=root2)
        else:
            cursor.execute("SELECT * from usuarios WHERE usuario=%s AND senha=%s", (usuario, senha,))
            result = cursor.fetchone()
            if result is None:
                    messagebox.showwarning('Login:', 'Usuário ou Senha inválidos.', parent=root2)
            else:
                global usuario_logado
                usuario_logado = result
                #print(usuario_logado)
                root2.destroy()
                lbl_logado1.configure(text= f' {usuario_logado[1]}')
                if usuario_logado[4] == 1:
                    nivel_acesso = 'Administrador'
                else:
                    nivel_acesso = 'Operador'
                lbl_logado2.configure(text= f'| {nivel_acesso}')
                setup_acesso()
                t = threading.Thread(target=bloqueio_botao)
                t.start()
                home()
                time.sleep(1)               
                r = threading.Thread(target=verifica_senha_protheus)
                r.start()
                time.sleep(1)
                s = threading.Thread(target=atualiza_op)
                s.start()                
                
    def logar_bind(event):
        logar()

    def sair():
        root2.destroy()
        root.destroy()
    #///////////////////////// LAYOUT
    frame0 = customtkinter.CTkFrame(root2, corner_radius=10, fg_color='#ffffff', border_width=4, border_color='#2a2d2e')
    frame0.pack(padx=4, pady=10, fill="both", expand=True)

    frame1 = Frame(frame0, bg='#ffffff')
    frame1.pack(padx=10, pady=5, fill="x", expand=False, side=TOP)
    frame2 = Frame(frame0, bg='#2a2d2e') #/// LINHA
    frame2.pack(padx=10, pady=0, fill="x", expand=False, side=TOP)
    frame3 = Frame(frame0, bg='#ffffff')
    frame3.pack(padx=10, pady=5, fill="x", expand=False, side=TOP)
    frame4 = Frame(frame0, bg='#2a2d2e') #/// LINHA
    frame4.pack(padx=10, pady=10, fill="x", expand=False, side=TOP)
    frame5 = Frame(frame0, bg='#ffffff')
    frame5.pack(padx=10, pady=5, fill="x", expand=False, side=TOP)


    #/////////FRAME1
    lbl_titulo = Label(frame1, text='Login', font=fonte_titulo_bold, bg='#ffffff', fg='#222222')
    lbl_titulo.grid(row=0, column=1)
    frame1.grid_columnconfigure(0, weight=1)
    frame1.grid_columnconfigure(2, weight=1)

    #/////////FRAME2 LINHA
        
    #/////////FRAME3
    lbl1=Label(frame3, text='Usuário:', font=fonte_padrao, bg='#ffffff', fg='#000000')
    lbl1.grid(row=0, column=1, sticky="w")
    ent_usuario = customtkinter.CTkEntry(frame3, **estilo_entry_padrao, width=282)
    ent_usuario.grid(row=1, column=1)
    ent_usuario.focus_force()
    ent_usuario.bind("<Return>", logar_bind)
   
    lbl2=Label(frame3, text='Senha:', font=fonte_padrao, bg='#ffffff', fg='#000000')
    lbl2.grid(row=2, column=1, sticky="w")
    ent_senha = customtkinter.CTkEntry(frame3, **estilo_entry_padrao, width=282, show='*')
    ent_senha.grid(row=3, column=1)
    ent_senha.bind("<Return>", logar_bind)

    frame3.grid_columnconfigure(0, weight=1)
    frame3.grid_columnconfigure(4, weight=1)

    #/////////FRAME4 LINHA

    #/////////FRAME5
    bt1 = customtkinter.CTkButton(frame5, text='Logar', **estilo_botao_padrao_form, width=134, command=logar)
    bt1.grid(row=0, column=1, padx=5)
    bt1 = customtkinter.CTkButton(frame5, text='Sair', **estilo_botao_padrao_form, width=134, command=sair)
    bt1.grid(row=0, column=2, padx=5)
    frame5.grid_columnconfigure(0, weight=1)
    frame5.grid_columnconfigure(3, weight=1)

    #ent_usuario.insert(0, 'adm_ivan.casagrande')
    #ent_senha.insert(0,'6176Ic12')

    #root2.update()
    #largura = root2.winfo_width()
    #altura = root2.winfo_height()
    #root2.wm_protocol("WM_DELETE_WINDOW", lambda: [home(), root2.destroy()])
    root2.mainloop()

def atualiza_senha_protheus():
    root2 = Toplevel(root)
    root2.bind_class("Button", "<Key-Return>", lambda event: event.widget.invoke())
    root2.unbind_class("Button", "<Key-space>")
    root2.focus_force()
    root2.grab_set()

    window_width = 360
    window_height = 320
    screen_width = root2.winfo_screenwidth()
    screen_height = root2.winfo_screenheight() - 70
    x_cordinate = int((screen_width / 2) - (window_width / 2))
    y_cordinate = int((screen_height / 2) - (window_height / 2))
    root2.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
    root2.resizable(0, 0)
    root2.configure(bg='#ffffff')
    root2.title(titulo)
    root2.overrideredirect(True)
    root2.iconbitmap('img\\icone.ico')
    root2.wm_protocol("WM_DELETE_WINDOW", lambda: [home(), root2.destroy()])    
    
    #///////////////////////// FUNÇÕES
    def atualizar():
        senha = ent_senha.get()
        if senha == "":
            messagebox.showwarning('Atualizar Senha:', 'Digite sua senha.', parent=root2)
        else:
            try:
                autent = usuario_logado[2]+':'+usuario_logado[3]
                autent_b = autent.encode()
                autorizacao = base64.b64encode(autent_b)
                #url = 'http://192.168.1.18:8689/rest/api/oauth2/v1/token' //Desenv
                url = 'http://192.168.1.18:8683/rest/api/oauth2/v1/token'                
                headers = {
                    'POST': '/rest/api/oauth2/v1/token',
                    'Host': 'http://192.168.1.18:8683',
                    'Accept': 'application/json',
                    'Authentication': 'BASIC '+ str(autorizacao),
                }
                parametros = {
                    'grant_type':'password',
                    'username':f'{usuario_logado[2]}',
                    'password':f'{senha}',
                }
                x = requests.post(url, headers=headers, params=parametros, timeout=2)
                x = x.status_code
                #print(x)
            except Exception as e:
                lbl_status_erro1.configure(text= f'Erro - Login(Função: atualiza_senha_protheus) / {e}')
                print('Erro - Login(Função: atualiza_senha_protheus) /', e)                
                pass

            if x == 200 or x == 201:
                try:
                    cursor.execute("UPDATE usuarios SET senha = %s where id = %s", (senha, usuario_logado[0]))
                    db.commit()
                except Exception as e:
                    messagebox.showerror('Atualizar Senha:', 'Erro de conexão com o banco de dados local.', parent=root2)
                    print(e)
                    return False
                messagebox.showinfo('Atualizar Senha:', 'Senha atualizada com sucesso.', parent=root2)
                root2.destroy()
                home()
            else:
                messagebox.showwarning('Atualizar Senha:', 'Senha incorreta.', parent=root2)

    def logar_bind(event):
        atualizar()

    def sair():
        root2.destroy()
        root.destroy()
    
    #///////////////////////// LAYOUT
    frame0 = customtkinter.CTkFrame(root2, corner_radius=10, fg_color='#ffffff', border_width=4, border_color='#2a2d2e')
    frame0.pack(padx=4, pady=10, fill="both", expand=True)

    frame1 = Frame(frame0, bg='#ffffff')
    frame1.pack(padx=10, pady=5, fill="x", expand=False, side=TOP)
    frame2 = Frame(frame0, bg='#2a2d2e') #/// LINHA
    frame2.pack(padx=10, pady=0, fill="x", expand=False, side=TOP)
    frame3 = Frame(frame0, bg='#ffffff')
    frame3.pack(padx=10, pady=5, fill="x", expand=False, side=TOP)
    frame4 = Frame(frame0, bg='#2a2d2e') #/// LINHA
    frame4.pack(padx=10, pady=10, fill="x", expand=False, side=TOP)
    frame5 = Frame(frame0, bg='#ffffff')
    frame5.pack(padx=10, pady=5, fill="x", expand=False, side=TOP)


    #/////////FRAME1
    lbl_titulo = Label(frame1, text='Atualização de Senha', font=fonte_titulo_bold, bg='#ffffff', fg='#222222')
    lbl_titulo.grid(row=0, column=1)
    frame1.grid_columnconfigure(0, weight=1)
    frame1.grid_columnconfigure(2, weight=1)

    #/////////FRAME2 LINHA
        
    #/////////FRAME3
    lbl1=Label(frame3, text='Atenção:\nSua senha encontra-se desatualizada\nem relação ao Sistema Protheus.', font=fonte_padrao, bg='#ffffff', fg='#000000')
    lbl1.grid(row=0, column=1, pady=(0,20))

    lbl2=Label(frame3, text='Digite sua nova senha:', font=fonte_padrao, bg='#ffffff', fg='#000000')
    lbl2.grid(row=1, column=1, sticky="w")
    ent_senha = customtkinter.CTkEntry(frame3, **estilo_entry_padrao, width=282, show='*')
    ent_senha.grid(row=2, column=1)
    ent_senha.bind("<Return>", logar_bind)

    frame3.grid_columnconfigure(0, weight=1)
    frame3.grid_columnconfigure(4, weight=1)

    #/////////FRAME4 LINHA

    #/////////FRAME5
    bt1 = customtkinter.CTkButton(frame5, text='Atualizar Senha', **estilo_botao_padrao_form, width=134, command=atualizar)
    bt1.grid(row=0, column=1, padx=5)
    bt1 = customtkinter.CTkButton(frame5, text='Sair', **estilo_botao_padrao_form, width=134, command=sair)
    bt1.grid(row=0, column=2, padx=5)
    frame5.grid_columnconfigure(0, weight=1)
    frame5.grid_columnconfigure(3, weight=1)

    #root2.update()
    #largura = root2.winfo_width()
    #altura = root2.winfo_height()

    root2.mainloop()

def sair():
    root.destroy()

def loop_status(): #Looping principal do sistema
    root.after(180000, status_protheus_thread)

def status_protheus_thread():
    u = threading.Thread(target=status_protheus)
    u.start()
    
def status_protheus(): #// Verifica se o serviço Rest está ativo ou não
    print('Função status_protheus')
    #url = 'http://192.168.1.18:8689/rest/' //Desenv
    url = 'http://192.168.1.18:8683/rest/'
    r = ''
    try:
        r = requests.get(url, timeout=0.9)
        r = r.status_code
        #print(r)
    except Exception as e:
        global status
        status = 1 # Se o status for 1, a conexão caiu. Lógica usada para verificar a sequencia ao abrir o programa. 
        lbl_status.configure(text= '| Desconectado ', image=nova_image_status2)
        lbl_status_erro1.configure(text= f'Erro - Protheus(Função: status_protheus) / {e}')        
        print('Erro - Função status_protheus /', e)
    if r == 200:
        lbl_status_erro1.configure(text='')
        lbl_status.configure(text= '| Conectado ', image=nova_image_status)
        if usuario_logado != '':
            s = threading.Thread(target=envia_protheus)
            s.start()
            #print('Enviou para protheus')

    root.after(0, loop_status)

def atualiza_op(): #// Atualiza a Ordem de Produção (OP) de acordo com o Protheus
    print('Função atualiza_op')
    r=''
    cod_retorno = ''
    try:
        url = 'http://192.168.1.18:8683/rest/SIMEC_PCP/op'
        r = requests.get(url, auth=(usuario_logado[2], usuario_logado[3]),timeout=6)
        cod_retorno = r.status_code
    except Exception as e:
        lbl_status_erro1.configure(text= f'Erro - Protheus(Função: atualiza_op) / {e}')        
        print('Erro - Função atualiza_op /', e)
        pass

    if cod_retorno == 200:
        lista_ops_protheus = []
        data = json.loads(r.content)
        for i in data:
            lista_ops_protheus.append(i['numop'])
            cursor.execute("select * from op where numop=%s",(i['numop'],))
            verifica = cursor.fetchone()
            if verifica == None:
                cursor.execute("INSERT INTO op (numop, produto, inmetro, tipo, inicio, termino, quant, desc1, desc2, desc3) values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",(i['numop'],i['produto'].rstrip(),i['inmetro'].rstrip(),i['tipo'], i['inicio'], i['termino'], i['quant'], i['desc1'].rstrip(), i['desc2'].rstrip(), i['desc3'].rstrip()))
        
        cursor.execute("select numop from op")
        lista_ops_local = cursor.fetchall()

        for i in lista_ops_local:
            if i[0] not in lista_ops_protheus:
                cursor.execute("update op set encerrado = 'OK' where numop = %s",(i[0],))
        db.commit()

def envia_protheus(): #//Envia para a tabela temporaria do Protheus
    print('Função envia_protheus')
    cursor.execute("SELECT * FROM simec_etiquetas.producao where envio_protheus is null or envio_protheus != 'OK' and envio_protheus != 'Processando'")
    if usuario_logado != '':
        lista = []
        cont = 0
        for i in cursor:
            dic = {}    
            if cont <10:
                if i[11] != 'Processando' and i[11] != 'OK':
                    peso = float(i[9])
                    resto = peso%1000
                    if resto == 0:
                        peso =int(str((peso/1000)).rstrip('.0'))
                    else:
                        peso = peso/1000
                    ano =(i[1])[6:]
                    mes =(i[1])[3:5]
                    dia =(i[1])[0:2]                    
                    data_b = ano+mes+dia
                    dic['op'] = i[10]
                    dic['corrida'] = i[7]
                    dic['volume'] = i[8]
                    dic['data'] = data_b
                    dic['operador'] = i[3]
                    dic['peso'] = peso
                    lista.append(dic)
                    cont += 1
            else:
                break
        retorno = ''
        try:
            url = 'http://192.168.1.18:8683/rest/SIMEC_PCP/producao'
            r = requests.post(url, json = lista, auth=(usuario_logado[2], usuario_logado[3]), timeout=9)
            retorno = json.loads(r.content)
            #print(retorno)
        except Exception as e:
            lbl_status_erro1.configure(text= f'Erro - Protheus(Função: envia_protheus) / {e}')                    
            print('Erro - Função envia_protheus /',e)
            pass

        if len(lista) != 0:            
            for i in retorno:
                if i['status'] == True:
                    cursor.execute("update producao set envio_protheus = %s, envio_data = %s, envio_hora = %s WHERE lote = %s",('Processando',data,hora,i['volume'],))
           
            db.commit()
            time.sleep(10)
        u = threading.Thread(target=envia_protheus_confirma)
        u.start()
        
def envia_protheus_confirma(): #//Confirma o registro no Protheus
    print('Função envia_protheus_confirma')
    cursor.execute("SELECT * FROM simec_etiquetas.producao where envio_protheus is null or envio_protheus != 'OK' and envio_tentativa <=3")
    if usuario_logado != '':
        lista_confirma = []
        cont = 0
        for i in cursor:
            dic = {}    
            if cont <10:
                if i[11] != 'OK' and i[14] <=3:
                    dic['volume'] = i[8]
                    lista_confirma.append(dic)
                    cont += 1
            else:
                break
        retorno = ''
        #print(lista_confirma)
        try:
            url = 'http://192.168.1.18:8683/rest/SIMEC_PCP/confapont'
            r = requests.get(url, json = lista_confirma, auth=(usuario_logado[2], usuario_logado[3]), timeout=6)
            retorno = json.loads(r.content)
            print(retorno, 'confirma')
        except Exception as e:
            lbl_status_erro1.configure(text= f'Erro - Protheus(Função: envia_protheus_confirma) / {e}')                    
            print('Erro - Função envia_protheus_confirma /', e)
            pass

        for i in retorno:
            if i['status'] == True:
                cursor.execute("update producao set envio_protheus = %s, envio_data = %s, envio_hora = %s WHERE lote = %s",('OK',data,hora,i['volume'],))
            elif i['status'] == False and i['erro'] != '':
                cursor.execute("select envio_tentativa from producao where lote = %s",(i['volume'],))
                erro = cursor.fetchone()[0]
                erro = erro + 1
                cursor.execute("update producao set envio_protheus = %s, envio_data = %s, envio_hora = %s, envio_tentativa = %s, erro_desc = %s WHERE lote = %s",('Erro',data,hora,erro,i['erro'],i['volume'],))
        db.commit()
        a = threading.Thread(target=envia_email)
        a.start()        
        time.sleep(10)
        u = threading.Thread(target=atualiza_op)
        u.start()

def envia_email(): #// Depois de 4 tentativas de enviar o lote pro Protheus, o sistema envia um email para o grupo simec.etiquetas@gruposimec.com.br
    print('Função envia_e-mail')    
    lista_email = []
    confirma_envio = 0
    try:
        cursor.execute("select * from producao where envio_protheus = 'Erro' and envio_tentativa >= 3")
        resultado_email = cursor.fetchall()
        #print(resultado_email)
        string_email = ''
        for i in resultado_email:
            if i[16] != 'OK':
                string_email = string_email + '<font color="#1D366C"><b>Lote: '+i[8]+'</b></font><br><b>Op:</b> '+ i[10]+'&emsp;<b>Data:</b> '+ i[1]+'&emsp; <b>Hora:</b> '+ i[2]+'&emsp; <b>Usuário:</b> '+ i[3]+'&emsp; <b>Produto:</b> '+ i[6]+'&emsp; <b>Corrida:</b> '+ i[7]+'&emsp; <b>Peso:</b> '+ i[9]+'&emsp; <b>Status:</b> '+ i[11]+'&emsp;<b>Data envio Protheus: </b>'+ i[12]+ '&emsp;<b>Hora envio Protheus: </b>'+ i[13]+'<br><b>Descrição do erro: </b><br>'+ i[15][:2500]+'<br><br><br><br>'
                lista_email.append(i)
        if len(lista_email) != 0:
            sender_email = "naoresponder@gruposimec.com.br"
            send_to=['simec.etiquetas@gruposimec.com.br','cesar@caerp.com.br']
            password = "Qu@@258147"
            toAddress = send_to

            message = MIMEMultipart("alternative")
            message["Subject"] = "Simec Etiquetas - Erro de Apontamento de Produção - (Não responder)"
            message["From"] = sender_email
            message["To"] = ','.join(send_to)

            # Create the plain-text and HTML version of your message
            text = """ """

            html = """\
            <html>
                <body>
                <center>	
            <font size="2" face="Arial" >
            <table width=100% border=0>
                <tr style="background-color:#01336e">
                <td align=center><p style= "font-family:Arial; font-size:50px; color:white"><b>Simec Etiquetas</b></p></td>
                </tr>
                <tr>
                <td align=center><b><h3><br>Atenção. Os seguintes lotes não foram apontados pelo sistema de etiquetas:</h3></b><br></td>
                </tr>
                <tr>
                <td>"""+str(string_email)+"""</td>
                </tr>
            </table>
            </font>
            </center>	
            </body>
            </html>

            """

            # Turn these into plain/html MIMEText objects
            part1 = MIMEText(text, "plain")
            part2 = MIMEText(html, "html")

            # Add HTML/plain-text parts to MIMEMultipart message
            # The email client will try to render the last part first
            message.attach(part1)
            message.attach(part2)

            # Create secure connection with server and send email
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtps.uhserver.com", 465, context=context) as server:
                server.login(sender_email, password)
                server.sendmail(
                    sender_email, toAddress, message.as_string()
                )
    except Exception as e:
        confirma_envio = 1
        lbl_status_erro1.configure(text= f'Erro - E-mail(Função: envia_email) / {e}')        
        print('Erro - Envia_email /', e)
        pass
    
    if confirma_envio == 0:
        for z in lista_email:
            cursor.execute("update producao set envio_email = 'OK' where id = %s",(z[0],))
        db.commit()

def verifica_sequencia(): #// Verifica a sequencia de lote no Protheus. Executado apenas uma vez quando abre o sistema. 
    time.sleep(3)
    print('Função verifica sequencia')
    try:
        url = 'http://192.168.1.18:8683/rest/SIMEC_PCP/seq'
        r = requests.get(url, auth=(usuario_logado[2], usuario_logado[3]), timeout=7)
        data = json.loads(r.content)
        print(data)
        cursor.execute("SELECT * from prod_tipo")
        resultado = cursor.fetchall()
        for i in resultado:
            if i[0] == 1:
                local_Barra = i[3]
            elif i[0] == 2:
                local_Rolo = i[3]    
            elif i[0] == 3:
                local_FioMaquina = i[3]    
            elif i[0] == 4:
                local_Endireitado = i[3]    
        
        if data != '':
            for i in data:
                if i['tipo'] == 'V':
                    if ord(local_Barra[1:2]) >= ord(i['sequencia'][1:2]): #//Verifica ano.
                        if ord(local_Barra[2:3]) >= ord(i['sequencia'][2:3]): #//Verifica mês
                            if int(local_Barra[3:]) < int(i['sequencia'][3:]): #//Verifica sequencia
                                cursor.execute("UPDATE prod_tipo SET ultima_seq = %s WHERE id = 1", (i['sequencia'],))
                                db.commit()
                                print('Sequencia da API é maior que a sequencia Local (Barra)')
                        else:
                            cursor.execute("UPDATE prod_tipo SET ultima_seq = %s WHERE id = 1", (i['sequencia'],))
                            db.commit()
                            print('Mês da API é maior que o mês Local (Barra)')                                
                    else:
                        cursor.execute("UPDATE prod_tipo SET ultima_seq = %s WHERE id = 1", (i['sequencia'],))
                        db.commit()
                        print('Ano da API é maior que o ano Local (Barra)')

                elif i['tipo'] == 'R':
                    if ord(local_Rolo[1:2]) >= ord(i['sequencia'][1:2]): #//Verifica ano.
                        if ord(local_Rolo[2:3]) >= ord(i['sequencia'][2:3]): #//Verifica mês
                            if int(local_Rolo[3:]) < int(i['sequencia'][3:]): #//Verifica sequencia
                                cursor.execute("UPDATE prod_tipo SET ultima_seq = %s WHERE id = 2", (i['sequencia'],))
                                db.commit()
                                print('Sequencia da API é maior que a sequencia Local (Rolo)')
                        else:
                            cursor.execute("UPDATE prod_tipo SET ultima_seq = %s WHERE id = 2", (i['sequencia'],))
                            db.commit()
                            print('Mês da API é maior que o mês Local (Rolo)')
                    else:
                        cursor.execute("UPDATE prod_tipo SET ultima_seq = %s WHERE id = 2", (i['sequencia'],))
                        db.commit()
                        print('Ano da API é maior que o ano Local (Rolo)')


                elif i['tipo'] == 'F':
                    if ord(local_FioMaquina[1:2]) >= ord(i['sequencia'][1:2]): #//Verifica ano.
                        if ord(local_FioMaquina[2:3]) >= ord(i['sequencia'][2:3]): #//Verifica mês
                            if int(local_FioMaquina[3:]) < int(i['sequencia'][3:]): #//Verifica sequencia
                                cursor.execute("UPDATE prod_tipo SET ultima_seq = %s WHERE id = 3", (i['sequencia'],))
                                db.commit()
                                print('Sequencia da API é maior que a sequencia Local (Fio Maquina)')
                        else:
                            cursor.execute("UPDATE prod_tipo SET ultima_seq = %s WHERE id = 3", (i['sequencia'],))
                            db.commit()
                            print('Mês da API é maior que o mês Local (Fio Maquina)')
                    else:
                        cursor.execute("UPDATE prod_tipo SET ultima_seq = %s WHERE id = 3", (i['sequencia'],))
                        db.commit()
                        print('Ano da API é maior que o ano Local (Fio Maquina)')

                elif i['tipo'] == 'E':
                    if ord(local_Endireitado[1:2]) >= ord(i['sequencia'][1:2]): #//Verifica ano.
                        if ord(local_Endireitado[2:3]) >= ord(i['sequencia'][2:3]): #//Verifica mês
                            if int(local_Endireitado[3:]) < int(i['sequencia'][3:]): #//Verifica sequencia
                                cursor.execute("UPDATE prod_tipo SET ultima_seq = %s WHERE id = 4", (i['sequencia'],))
                                db.commit()
                                print('Sequencia da API é maior que a sequencia Local (Endireitado)')
                        else:
                            cursor.execute("UPDATE prod_tipo SET ultima_seq = %s WHERE id = 4", (i['sequencia'],))
                            db.commit()
                            print('Mês da API é maior que o mês Local (Endireitado)')
                    else:
                        cursor.execute("UPDATE prod_tipo SET ultima_seq = %s WHERE id = 4", (i['sequencia'],))
                        db.commit()
                        print('Ano da API é maior que o ano Local (Endireitado)')
            
    except Exception as e:
        if usuario_logado != '':
            lbl_status_erro1.configure(text= f'Erro - Protheus(Função: verifica_sequencia) / {e}')        
            print('Erro - Função verifica sequencia /', e)
        if status == 0: # Se o status for 0 a conexão está ok. Se entrou aqui é pq a API teve algum problema e vai ser executada novamente. 
            h = threading.Thread(target=verifica_sequencia)
            h.start()
        else:
            pass

root=customtkinter.CTk()
root.state('zoomed')
root.title(titulo)
root.after(0, login)
root.iconbitmap('img\\icone.ico')

frame1 = Frame(root, bg="#dcdcdc", height=20)
frame1.pack(side=TOP, fill=X, expand=False)
frame2 = Frame(root, bg="#222222", height=20)
frame2.pack(side=BOTTOM, fill=X, expand=False)
frame3 = Frame(root, bg="#1D366C")
frame3.pack(side=LEFT, fill=Y, expand=False)
frame4 = Frame(root, bg="#ffffff")
frame4.pack(side=RIGHT, fill=BOTH, expand=True)

image_logo = Image.open('img\\logo.png')
resize_logo = image_logo.resize((60, 60))
nova_image_logo = ImageTk.PhotoImage(resize_logo)
lbl_logado = Label(frame1, text="SIMEC ETIQUETAS", image=nova_image_logo, compound="left", bg='#dcdcdc', fg="#1D366C", font=fonte_titulo_bold)
lbl_logado.grid(row=0, column=0, padx=10)

frame1.grid_columnconfigure(1, weight=1)
frame1.grid_columnconfigure(1, weight=1)    

image_logado = Image.open('img\\logado.png')
resize_logado = image_logado.resize((30, 30))
nova_image_logado = ImageTk.PhotoImage(resize_logado)
lbl_logado1 = Label(frame1, text=" ", image=nova_image_logado, compound="left", bg='#dcdcdc', fg="#1D366C", font=fonte_padrao_bold)
lbl_logado1.grid(row=0, column=2)

lbl_logado2 = Label(frame1, text="", bg='#dcdcdc', fg="#1D366C", font=fonte_padrao_bold)
lbl_logado2.grid(row=0, column=3, padx=(0,0))

image_status2 = Image.open('img\\vermelho.png')
resize_status2 = image_status2.resize((20, 20))
nova_image_status2 = ImageTk.PhotoImage(resize_status2)

image_status = Image.open('img\\verde.png')
resize_status = image_status.resize((20, 20))
nova_image_status = ImageTk.PhotoImage(resize_status)
lbl_status = Label(frame1, text="", image=nova_image_status, compound="right", bg='#dcdcdc', fg="#1D366C", font=fonte_padrao_bold)
lbl_status.grid(row=0, column=4, padx=(0,10))

image_home = Image.open('img\\home.png')
resize_home = image_home.resize((30, 30))
nova_image_home = ImageTk.PhotoImage(resize_home)
btn_home = Button(frame3, image=nova_image_home, compound="top", text='Home', bg='#1D366C', fg='#FFFFFF',
                    activebackground='#1D366C', borderwidth=0,
                    activeforeground="#01a8f7", highlightthickness=4, relief=RIDGE, command=home,
                    font=fonte_botoes, cursor="hand2")
btn_home.grid(row=1, column=1, padx=5, pady=5)


image_login = Image.open('img\\login.png')
resize_login = image_login.resize((30, 30))
nova_image_login = ImageTk.PhotoImage(resize_login)
btn_login = Button(frame3, image=nova_image_login, compound="top", text='Login', bg='#1D366C', fg='#FFFFFF',
                    activebackground='#1D366C', borderwidth=0,
                    activeforeground="#01a8f7", highlightthickness=4, relief=RIDGE, command=login,
                    font=fonte_botoes, cursor="hand2")
btn_login.grid(row=2, column=1, padx=5, pady=5)

image_corrida = Image.open('img\\corrida.png')
resize_corrida = image_corrida.resize((30, 30))
nova_image_corrida = ImageTk.PhotoImage(resize_corrida)
btn_corrida = Button(frame3, image=nova_image_corrida, compound="top", text='Corridas', bg='#1D366C', fg='#FFFFFF',
                    activebackground='#1D366C', borderwidth=0,
                    activeforeground="#01a8f7", highlightthickness=4, relief=RIDGE, command=corridas,
                    font=fonte_botoes, cursor="hand2")
btn_corrida.grid(row=3, column=1, padx=5, pady=5)


image_etiqueta = Image.open('img\\impressora_menu.png')
resize_etiqueta = image_etiqueta.resize((50, 30))
nova_image_etiqueta = ImageTk.PhotoImage(resize_etiqueta)
btn_etiqueta = Button(frame3, image=nova_image_etiqueta, compound="top", text='Etiquetas', bg='#1D366C', fg='#FFFFFF',
                    activebackground='#1D366C', borderwidth=0,
                    activeforeground="#01a8f7", highlightthickness=4, relief=RIDGE, command=etiquetas,
                    font=fonte_botoes, cursor="hand2")
btn_etiqueta.grid(row=4, column=1, padx=5, pady=5)

image_reimp = Image.open('img\\reimpressao.png')
resize_reimp = image_reimp.resize((40, 30))
nova_image_reimp = ImageTk.PhotoImage(resize_reimp)
btn_reimp = Button(frame3, image=nova_image_reimp, compound="top", text='Reimpressão', bg='#1D366C', fg='#FFFFFF',
                    activebackground='#1D366C', borderwidth=0,
                    activeforeground="#01a8f7", highlightthickness=4, relief=RIDGE, command=reimp,
                    font=fonte_botoes, cursor="hand2")
btn_reimp.grid(row=5, column=1, padx=5, pady=5)

image_avulsa = Image.open('img\\avulsa.png')
resize_avulsa = image_avulsa.resize((30, 30))
nova_image_avulsa = ImageTk.PhotoImage(resize_avulsa)
btn_avulsa = Button(frame3, image=nova_image_avulsa, compound="top", text='Avulsas', bg='#1D366C', fg='#FFFFFF',
                    activebackground='#1D366C', borderwidth=0,
                    activeforeground="#01a8f7", highlightthickness=4, relief=RIDGE, command=avulsas,
                    font=fonte_botoes, cursor="hand2")
btn_avulsa.grid(row=6, column=1, padx=5, pady=5)

image_usuario = Image.open('img\\usuarios.png')
resize_usuario = image_usuario.resize((30, 30))
nova_image_usuario = ImageTk.PhotoImage(resize_usuario)
btn_usuario = Button(frame3, image=nova_image_usuario, compound="top", text='Usuários', bg='#1D366C', fg='#FFFFFF',
                    activebackground='#1D366C', borderwidth=0,
                    activeforeground="#01a8f7", highlightthickness=4, relief=RIDGE, command=cadastro_usuarios,
                    font=fonte_botoes, cursor="hand2")
btn_usuario.grid(row=8, column=1, padx=5, pady=5)

frame3.grid_rowconfigure(9, weight=1)
frame3.grid_rowconfigure(9, weight=1)

image_sair = Image.open('img\\sair.png')
resize_sair = image_sair.resize((25, 30))
nova_image_sair = ImageTk.PhotoImage(resize_sair)
btn_sair = Button(frame3, image=nova_image_sair, compound="top", text='Sair', bg='#1D366C', fg='#FFFFFF',
                    activebackground='#1D366C', borderwidth=0,
                    activeforeground="#01a8f7", highlightthickness=4, relief=RIDGE, command=sair,
                    font=fonte_botoes, cursor="hand2")
btn_sair.grid(row=10, column=1, padx=5, pady=5)

lbl_status_erro0 = Label(frame2, text= ' Último log de erro:', font=fonte_padrao, fg='#ffffff', background='#222222')
lbl_status_erro0.grid(row=0, column=1)

lbl_status_erro1 = Label(frame2, text= '', font=fonte_padrao, fg='#ffffff', background='#222222')
lbl_status_erro1.grid(row=0, column=2)

frame2.grid_rowconfigure(0, weight=1)
frame2.grid_rowconfigure(2, weight=1)

#/////////////////////////////BANCO DE DADOS/////////////////////////////
db = mysql.connector.connect(
    host="localhost",
    #host="192.168.11.125",
    #user="root",
    user="acesso_rede",
    passwd="senha",
    database="simec_etiquetas",
    connect_timeout=1000,
)
cursor = db.cursor(buffered=True)
#/////////////////////////////FIM BANCO DE DADOS/////////////////////////////

t = threading.Thread(target=status_protheus)
t.start()

s = threading.Thread(target=verifica_sequencia)
s.start()

root.mainloop()

