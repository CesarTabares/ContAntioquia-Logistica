# -*- coding: utf-8 -*-
"""
Created on Sat Feb  1 07:28:37 2020

@author: Cesar
"""

from datetime import datetime

import wx
import openpyxl

col_requerimiento_auto=1
col_fecha_auto=2
col_cotizacion=3
col_tipotransporte=4
col_tipocontenedor=5
col_requieredescargue=6
col_origen=7
col_destino=8
col_km=9
col_precio=10
col_recargopeaje=11
col_nombreresponsable=12
col_telefono_resp=13
col_cargo=14
col_nombresiso=15
col_telefono_siso=16
col_debeinfo=17
col_horasantes=18



principal_color=wx.Colour(51, 102, 51)
wb_listas=openpyxl.load_workbook('Listas.xlsx')
wb_req=openpyxl.load_workbook('db_req.xlsx')

class MyFrame(wx.Frame):
    
    
    def OnKeyDown(self, event):
        """quit if user press q or Esc"""
        if event.GetKeyCode() == 27 or event.GetKeyCode() == ord('Q'): #27 is Esc
            self.Close(force=True)
            
        else:
            event.Skip()
    
    def __init__(self):
        
        wx.Frame.__init__(self, None, wx.ID_ANY, "Contenedores de Antioquia - Centro Logistico", size=(800, 500))  
        self.Bind(wx.EVT_KEY_UP, self.OnKeyDown)
        
        try:
            
            #image_file = 'CINCO CONSULTORES.jpg'
            #bmp1 = wx.Image(
                #image_file, 
                #wx.BITMAP_TYPE_ANY).ConvertToBitmap()
            
            #self.panel = wx.StaticBitmap(
                #self, -1, bmp1, (0, 0)
            self.panel=wx.Panel(self)
            self.panel.SetBackgroundColour(principal_color)

        except IOError:
            print ("Image file %s not found"  )
            raise SystemExit
        
        ico = wx.Icon('Cont.ico', wx.BITMAP_TYPE_ICO)
        self.SetIcon(ico)
        self.fgs= wx.GridBagSizer(0,0)
        
        title_font= wx.Font(20, wx.FONTFAMILY_DECORATIVE, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_LIGHT)
       
        self.lbltitle =wx.StaticText(self.panel, label='CONTENEDORES DE ANTIOQUIA')
        self.lbltitle.SetFont(title_font)
        self.lbltitle.SetBackgroundColour(principal_color)
        self.lbltitle.SetForegroundColour('white')
        self.fgs.Add(self.lbltitle,pos=(2,3),span=(1,3), flag=wx.ALL | wx.ALIGN_CENTER, border=5)
            
        btn_nuevo_req = wx.Button(self.panel, id=wx.ID_ANY, label="NUEVO\nREQUERIMIENTO", size=(150,60))
        self.fgs.Add(btn_nuevo_req, pos=(11,2),span=(1,2), flag= wx.ALL| wx.ALIGN_CENTER, border=0)
        btn_nuevo_req.Bind(wx.EVT_BUTTON, self.open_nuevo_req11)
        
        btn_logistico = wx.Button(self.panel, id=wx.ID_ANY, label="LOGISTICA",size=(150,60))
        self.fgs.Add(btn_logistico, pos=(11,5),span=(1,2), flag= wx.ALL | wx.ALIGN_RIGHT, border=0)
        btn_logistico.Bind(wx.EVT_BUTTON, self.logistica)
        
        btn_logistico = wx.Button(self.panel, id=wx.ID_ANY, label="Configuracion",size=(-1,-1))
        self.fgs.Add(btn_logistico, pos=(17,6),span=(1,1), flag= wx.ALL, border=0)
        btn_logistico.Bind(wx.EVT_BUTTON, self.configuracion)
        
        
        
        mainSizer= wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(self.fgs,0, flag=wx.ALIGN_CENTER)
        self.panel.SetSizerAndFit(mainSizer)
            
    #-------------Button Functions-----------------#
    def open_nuevo_req11(self, event):
        ww_nuevo_requerimiento11(parent=self.panel).Show()
       

    def logistica(self, event): 
        #consultawindow=ww_Consultar_Proceso(parent=self.panel)
        #consultawindow.Show()
        pass
        
    def configuracion(self, event):
        print ("Button pressed!")

    #-------------Button Functions-----------------#              
    
class ww_nuevo_requerimiento11(wx.Frame):
    
    def __init__(self,parent):
        ######----------------------------------------BACK END----------------------------------------#############        
        
        req1_sheet=wb_listas['Requerimientos-1']
        
        areas=[]
        
        for cell in req1_sheet['A']:
            if cell.value != None:
                areas.append(cell.value)
        areas.pop(0)
                
        
        ######----------------------------------------BACK END----------------------------------------#############       
        
        ######----------------------------------------FRONT END----------------------------------------#############
        
        wx.Frame.__init__(self, None, wx.ID_ANY, "Contenedores de Antioquia - Centro Logistico", size=(250, 250))  
        
        try:
            
            #image_file = 'CINCO CONSULTORES.jpg'
            #bmp1 = wx.Image(
                #image_file, 
                #wx.BITMAP_TYPE_ANY).ConvertToBitmap()
            
            #self.panel = wx.StaticBitmap(
                #self, -1, bmp1, (0, 0)
            self.panel=wx.Panel(self)
            self.panel.SetBackgroundColour(principal_color)

        except IOError:
            print ("Image file %s not found"  )
            raise SystemExit
        
        ico = wx.Icon('Cont.ico', wx.BITMAP_TYPE_ICO)
        self.SetIcon(ico)
        self.fgs= wx.GridBagSizer(0,0)
        
        title_font= wx.Font(10, wx.FONTFAMILY_DECORATIVE, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
       
        self.lbltitle =wx.StaticText(self.panel, label='Nuevo Requerimiento Por:')
        self.lbltitle.SetFont(title_font)
        self.lbltitle.SetBackgroundColour(principal_color)
        self.lbltitle.SetForegroundColour('white')
        self.fgs.Add(self.lbltitle,pos=(2,1),span=(1,3), flag=wx.ALL | wx.ALIGN_CENTER, border=5)

        self.combo_area = wx.ComboBox(self.panel,value=areas[0], choices=areas)
        self.fgs.Add(self.combo_area , pos=(4,1),span=(1,3), flag= wx.ALL | wx.EXPAND, border=5)
        
        btn_aceptar = wx.Button(self.panel, id=wx.ID_ANY, label="Aceptar",size=(-1,-1))
        self.fgs.Add(btn_aceptar, pos=(6,1),span=(1,3), flag= wx.ALL | wx.ALIGN_CENTER, border=0)
        btn_aceptar.Bind(wx.EVT_BUTTON, self.open_nuevo_req12)

        mainSizer= wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(self.fgs,0, flag=wx.ALIGN_CENTER)
        self.panel.SetSizerAndFit(mainSizer)
        
    #-------------Button Functions-----------------#            
    def open_nuevo_req12(self, event):
        self.Destroy()
        ww_nuevo_requerimiento12(parent=self.panel).Show()

    #-------------Button Functions-----------------# 
        
#############----------------------------------------FRONT END----------------------------------------#############

class ww_nuevo_requerimiento12(wx.Frame):
    
    def __init__(self,parent):
        ######----------------------------------------BACK END----------------------------------------#############
        req2_sheet=wb_listas['Requerimientos-12']
        hist_req_sheet=wb_req['Requerimientos']
        
        lista_encargado=[]
        lista_tipo_cont=[]
        lista_tipo_transp=[]
        lista_descargue=[]
        lista_debe_enviarinfo=[]
        
        for cell in req2_sheet['A']:
            if cell.value != None:
                lista_encargado.append(cell.value)
        for cell in req2_sheet['B']:
            if cell.value != None:
                lista_tipo_cont.append(cell.value)
        for cell in req2_sheet['C']:
            if cell.value != None:
                lista_tipo_transp.append(cell.value)
        for cell in req2_sheet['D']:
            if cell.value != None:
                lista_descargue.append(cell.value)   
        for cell in req2_sheet['E']:
            if cell.value != None:
                lista_debe_enviarinfo.append(cell.value)
        
        fila_vacia = 1
        
        while (hist_req_sheet.cell(row = fila_vacia, column = 1).value != None) :
          fila_vacia += 1
        
        nro_req=fila_vacia-1
        
        
        ######----------------------------------------BACK END----------------------------------------#############       
        
        ######----------------------------------------FRONT END----------------------------------------#############
        
        wx.Frame.__init__(self, None, wx.ID_ANY, "Contenedores de Antioquia - Centro Logistico", size=(900, 700))  
        
        try:
            #image_file = 'CINCO CONSULTORES.jpg'
            #bmp1 = wx.Image(
                #image_file, 
                #wx.BITMAP_TYPE_ANY).ConvertToBitmap()
            
            #self.panel = wx.StaticBitmap(
                #self, -1, bmp1, (0, 0)
            self.panel=wx.Panel(self)
            self.panel.SetBackgroundColour(principal_color)

        except IOError:
            print ("Image file %s not found"  )
            raise SystemExit
        
        ico = wx.Icon('Cont.ico', wx.BITMAP_TYPE_ICO)
        self.SetIcon(ico)
        self.fgs= wx.GridBagSizer(0,0)
        title_font= wx.Font(10, wx.FONTFAMILY_DECORATIVE, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        
        self.lbltitle1 =wx.StaticText(self.panel, label='CN Contenedores de Antioquia')
        self.lbltitle2 =wx.StaticText(self.panel, label='CENTRO LOGISTICO')
        self.lblrequerimiento =wx.StaticText(self.panel, label='Requerimiento N°')
        self.lblrequerimiento_auto =wx.StaticText(self.panel, label=str(nro_req))
        self.lblfecha =wx.StaticText(self.panel, label='Fecha')
        self.lblfecha_auto =wx.StaticText(self.panel, label=datetime.today().strftime('%Y-%m-%d')) #-%H:%M:%S
        self.lblcotizacion =wx.StaticText(self.panel, label='Cotizacion N°')
        self.lbltipotransporte =wx.StaticText(self.panel, label='Tipo de Trnasporte')
        self.lbltipocontenedor =wx.StaticText(self.panel, label='Tipo de Contenedor')
        self.lblrequieredescargue =wx.StaticText(self.panel, label='Requiere Descargue')
        self.lblorigen =wx.StaticText(self.panel, label='Origen')
        self.lbldestino =wx.StaticText(self.panel, label='Destino')
        self.lblkm =wx.StaticText(self.panel, label='Km')
        self.lblprecio =wx.StaticText(self.panel, label='Precio')
        self.lblrecargopeaje =wx.StaticText(self.panel, label='Recargo\nPeaje')
        self.lblinfocliente =wx.StaticText(self.panel, label='INFORMACION CLIENTE')
        self.lblnombreresponsable =wx.StaticText(self.panel, label='Nombre Responsable')
        self.lbltelefono_resp =wx.StaticText(self.panel, label='Telefono')
        self.lblcargo =wx.StaticText(self.panel, label='Cargo')
        self.lblnombresiso =wx.StaticText(self.panel, label='Nombre SISO')
        self.lbltelefono_siso =wx.StaticText(self.panel, label='Telefono')
        self.lbldebeinfo =wx.StaticText(self.panel, label='Debe Enviarse\n Informacion')
        self.lblhorasantes =wx.StaticText(self.panel, label='N° Horas Antes')
        #self.lbl =wx.StaticText(self.panel, label='')
        self.lbltitle2.SetFont(title_font)
        self.lblinfocliente.SetFont(title_font)
        
        
        
        self.lbltitle1.SetBackgroundColour(principal_color)
        self.lbltitle2.SetBackgroundColour(principal_color)
        self.lblrequerimiento.SetBackgroundColour(principal_color)
        self.lblrequerimiento_auto.SetBackgroundColour(principal_color)
        self.lblfecha.SetBackgroundColour(principal_color)
        self.lblfecha_auto.SetBackgroundColour(principal_color)
        self.lblcotizacion.SetBackgroundColour(principal_color)
        self.lbltipotransporte.SetBackgroundColour(principal_color)
        self.lbltipocontenedor.SetBackgroundColour(principal_color)
        self.lblrequieredescargue.SetBackgroundColour(principal_color)
        self.lblorigen.SetBackgroundColour(principal_color)
        self.lbldestino.SetBackgroundColour(principal_color)
        self.lblkm.SetBackgroundColour(principal_color)
        self.lblprecio.SetBackgroundColour(principal_color)
        self.lblrecargopeaje.SetBackgroundColour(principal_color)
        self.lblinfocliente.SetBackgroundColour(principal_color)
        self.lblnombreresponsable.SetBackgroundColour(principal_color)
        self.lbltelefono_resp.SetBackgroundColour(principal_color)
        self.lblcargo.SetBackgroundColour(principal_color)
        self.lblnombresiso.SetBackgroundColour(principal_color)
        self.lbltelefono_siso.SetBackgroundColour(principal_color)
        self.lbldebeinfo.SetBackgroundColour(principal_color)
        self.lblhorasantes.SetBackgroundColour(principal_color)
        
        self.lbltitle1.SetForegroundColour('white')
        self.lbltitle2.SetForegroundColour('white')
        self.lblrequerimiento.SetForegroundColour('white')
        self.lblrequerimiento_auto.SetForegroundColour('white')
        self.lblfecha.SetForegroundColour('white')
        self.lblfecha_auto.SetForegroundColour('white')
        self.lblcotizacion.SetForegroundColour('white')
        self.lbltipotransporte.SetForegroundColour('white')
        self.lbltipocontenedor.SetForegroundColour('white')
        self.lblrequieredescargue.SetForegroundColour('white')
        self.lblorigen.SetForegroundColour('white')
        self.lbldestino.SetForegroundColour('white')
        self.lblkm.SetForegroundColour('white')
        self.lblprecio.SetForegroundColour('white')
        self.lblrecargopeaje.SetForegroundColour('white')
        self.lblinfocliente.SetForegroundColour('white')
        self.lblnombreresponsable.SetForegroundColour('white')
        self.lbltelefono_resp.SetForegroundColour('white')
        self.lblcargo.SetForegroundColour('white')
        self.lblnombresiso.SetForegroundColour('white')
        self.lbltelefono_siso.SetForegroundColour('white')
        self.lbldebeinfo.SetForegroundColour('white')
        self.lblhorasantes.SetForegroundColour('white')
        #self.lbl =wx.StaticText(self.panel, label='')
        
        self.txtcotizacion=wx.TextCtrl(self.panel)
        self.txtorigen=wx.TextCtrl(self.panel)
        self.txtdestino=wx.TextCtrl(self.panel)
        self.txtkm=wx.TextCtrl(self.panel)
        self.txtprecio=wx.TextCtrl(self.panel)
        self.txtnombreresponsable=wx.TextCtrl(self.panel)
        self.txttelefono_resp=wx.TextCtrl(self.panel)
        self.txtcargo=wx.TextCtrl(self.panel)
        self.txtnombresiso=wx.TextCtrl(self.panel)
        self.txttelefono_siso=wx.TextCtrl(self.panel)
        self.txthorasantes=wx.TextCtrl(self.panel)
        
        self.combotipotransporte=wx.ComboBox(self.panel,value=lista_tipo_transp[0], choices=lista_tipo_transp)
        self.combotipocontenedor=wx.ComboBox(self.panel,value=lista_tipo_cont[0], choices=lista_tipo_cont)
        self.comborequieredescargue=wx.ComboBox(self.panel,value=lista_descargue[0], choices=lista_descargue)
        self.combodebeinfo=wx.ComboBox(self.panel,value=lista_debe_enviarinfo[0], choices=lista_debe_enviarinfo)
        
        btn_guardar = wx.Button(self.panel, id=wx.ID_ANY, label="Guardar",size=(-1,-1))
        btn_salir = wx.Button(self.panel, id=wx.ID_ANY, label="Salir",size=(-1,-1))
        btn_adicionar_transp = wx.Button(self.panel, id=wx.ID_ANY, label="Adicionar",size=(-1,-1))
        
        self.fgs.Add(btn_adicionar_transp, pos=(15,8),span=(1,1), flag= wx.ALL | wx.ALIGN_CENTER, border=5)
        self.fgs.Add(btn_guardar, pos=(16,8),span=(1,1), flag= wx.ALL | wx.ALIGN_CENTER, border=5)
        self.fgs.Add(btn_salir, pos=(17,8),span=(1,1), flag= wx.ALL | wx.ALIGN_CENTER, border=5)
               
        self.fgs.Add(self.combotipotransporte,pos=(6,2),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.combotipocontenedor,pos=(7,2),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.comborequieredescargue,pos=(8,2),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.combodebeinfo,pos=(15,2),span=(1,1), flag= wx.ALL, border=5)

        self.fgs.Add(self.txtcotizacion, pos=(4,2),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.txtorigen, pos=(6,4),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.txtdestino, pos=(6,5),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.txtkm, pos=(6,6),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.txtprecio, pos=(8,5),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.txtnombreresponsable, pos=(12,2),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.txttelefono_resp, pos=(12,5),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.txtcargo, pos=(12,8),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.txtnombresiso, pos=(14,2),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.txttelefono_siso, pos=(14,5),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.txthorasantes, pos=(16,2),span=(1,1), flag= wx.ALL, border=5)

        self.fgs.Add(self.lbltitle1 , pos=(1,1),span=(1,2), flag= wx.ALL, border=5)
        self.fgs.Add(self.lbltitle2 , pos=(1,3),span=(1,4), flag= wx.ALL | wx.ALIGN_CENTER, border=5)
        self.fgs.Add(self.lblrequerimiento , pos=(2,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrequerimiento_auto, pos=(2,2),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblfecha , pos=(2,7),span=(1,1), flag= wx.ALL | wx.ALIGN_RIGHT, border=5)
        self.fgs.Add(self.lblfecha_auto , pos=(2,8),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblcotizacion , pos=(4,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lbltipotransporte , pos=(6,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lbltipocontenedor , pos=(7,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrequieredescargue, pos=(8,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblorigen , pos=(5,4),span=(1,1), flag= wx.ALL | wx.ALIGN_CENTER, border=0)
        self.fgs.Add(self.lbldestino , pos=(5,5),span=(1,1), flag= wx.ALL| wx.ALIGN_CENTER, border=0)
        self.fgs.Add(self.lblkm , pos=(5,6),span=(1,1), flag= wx.ALL| wx.ALIGN_CENTER, border=0)
        self.fgs.Add(self.lblprecio , pos=(7,5),span=(1,1), flag= wx.ALL |wx.ALIGN_BOTTOM | wx.ALIGN_CENTER_HORIZONTAL, border=0)
        self.fgs.Add(self.lblrecargopeaje , pos=(7,6),span=(2,1), flag= wx.ALL| wx.ALIGN_CENTER, border=5)
        self.fgs.Add(self.lblinfocliente , pos=(10,4),span=(1,3), flag= wx.ALL| wx.ALIGN_CENTER, border=5)
        self.fgs.Add(self.lblnombreresponsable , pos=(12,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lbltelefono_resp , pos=(12,4),span=(1,1), flag= wx.ALL | wx.ALIGN_RIGHT, border=5)
        self.fgs.Add(self.lblcargo , pos=(12,7),span=(1,1), flag= wx.ALL| wx.ALIGN_RIGHT, border=5)
        self.fgs.Add(self.lblnombresiso , pos=(14,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lbltelefono_siso , pos=(14,4),span=(1,1), flag= wx.ALL |wx.ALIGN_RIGHT, border=5)
        self.fgs.Add(self.lbldebeinfo , pos=(15,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblhorasantes , pos=(16,1),span=(1,1), flag= wx.ALL, border=5)
        
        
        btn_guardar.Bind(wx.EVT_BUTTON, self.guardar_req)
        btn_salir.Bind(wx.EVT_BUTTON, self.salir)
        btn_adicionar_transp.Bind(wx.EVT_BUTTON, self.adicionar_transp)
        
        mainSizer= wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(self.fgs,0, flag=wx.ALIGN_LEFT)
        self.panel.SetSizerAndFit(mainSizer)
        
    def guardar_req(self,event):
        
        hist_req_sheet=wb_req['Requerimientos']
        
        fila_vacia = 1
        
        while (hist_req_sheet.cell(row = fila_vacia, column = 1).value != None) :
          fila_vacia += 1
        
        nro_req=fila_vacia-1
        
        requerimiento_auto=self.lblrequerimiento_auto.GetLabel()
        fecha_auto=self.lblfecha_auto.GetLabel()
        cotizacion=self.txtcotizacion.GetValue()
        tipotransporte=self.combotipotransporte.GetValue()
        tipocontenedor=self.combotipocontenedor.GetValue()
        requieredescargue=self.comborequieredescargue.GetValue()
        origen=self.txtorigen.GetValue()
        destino=self.txtdestino.GetValue()
        km=self.txtkm.GetValue()
        precio=self.txtprecio.GetValue()
        nombreresponsable=self.txtnombreresponsable.GetValue()
        telefono_resp=self.txttelefono_resp.GetValue()
        cargo=self.txtcargo.GetValue()
        nombresiso=self.txtnombresiso.GetValue()
        telefono_siso=self.txttelefono_siso.GetValue()
        debeinfo=self.combodebeinfo.GetValue()
        horasantes=self.txthorasantes.GetValue()
        
        hist_req_sheet.cell(row=fila_vacia, column=col_requerimiento_auto).value=requerimiento_auto
        hist_req_sheet.cell(row=fila_vacia, column=col_fecha_auto).value=fecha_auto
        hist_req_sheet.cell(row=fila_vacia, column=col_cotizacion).value=cotizacion
        hist_req_sheet.cell(row=fila_vacia, column=col_tipotransporte).value=tipotransporte
        hist_req_sheet.cell(row=fila_vacia, column=col_tipocontenedor).value=tipocontenedor
        hist_req_sheet.cell(row=fila_vacia, column=col_requieredescargue).value=requieredescargue
        hist_req_sheet.cell(row=fila_vacia, column=col_origen).value=origen
        hist_req_sheet.cell(row=fila_vacia, column=col_destino).value=destino
        hist_req_sheet.cell(row=fila_vacia, column=col_km).value=km
        hist_req_sheet.cell(row=fila_vacia, column=col_precio).value=precio
        hist_req_sheet.cell(row=fila_vacia, column=col_nombreresponsable).value=nombreresponsable
        hist_req_sheet.cell(row=fila_vacia, column=col_telefono_resp).value=telefono_resp
        hist_req_sheet.cell(row=fila_vacia, column=col_cargo).value=cargo
        hist_req_sheet.cell(row=fila_vacia, column=col_nombresiso).value=nombresiso
        hist_req_sheet.cell(row=fila_vacia, column=col_telefono_siso).value=telefono_siso
        hist_req_sheet.cell(row=fila_vacia, column=col_debeinfo).value=debeinfo
        hist_req_sheet.cell(row=fila_vacia, column=col_horasantes).value=horasantes
        
        self.combotipotransporte.Value=''
        self.combotipocontenedor.Value=''
        self.comborequieredescargue.Value=''
        self.txtorigen.Value=''
        self.txtdestino.Value=''
        self.txtkm.Value=''
        self.txtprecio.Value=''
        self.txtnombreresponsable.Value=''
        self.txttelefono_resp.Value=''
        self.txtcargo.Value=''
        self.txtnombresiso.Value=''
        self.txttelefono_siso.Value=''
        self.combodebeinfo.Value=''
        self.txthorasantes.Value=''

        wb_req.save('db_req.xlsx')

    def salir(self,event):
        pass
    
    def adicionar_transp(self,event):
        pass
        
class MyApp(wx.App):
    def OnInit(self):
        self.frame= MyFrame()
        self.frame.Show()
        return True       
 
# Run the program     
app=MyApp()
app.MainLoop()
del app
            
            