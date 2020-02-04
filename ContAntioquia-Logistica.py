# -*- coding: utf-8 -*-
"""
Created on Sat Feb  1 07:28:37 2020

@author: Cesar
"""

from datetime import datetime

import wx
import openpyxl

from pubsub import pub

col_requerimiento_auto=1
col_area_req=2
col_area=6
col_fecha_auto=3
col_cotizacion=4
col_tipotransporte=5
col_tipocontenedor=7
col_requieredescargue=8
col_origen=9
col_destino=10
col_km=11
col_precio=12
col_recargopeaje=13
col_nombreresponsable=14
col_telefono_resp=15
col_cargo=16
col_nombresiso=17
col_telefono_siso=18
col_debeinfo=19
col_horasantes=20
col_fechaentrega=21
col_direccion=22
col_referenciacont=23
col_nombreconduc=24
col_cedula=25
col_telefonoconduc=26
col_placa=27
col_adiciones=28
col_preguntahoras=29
col_preguntadoc=30




principal_color=wx.Colour(51, 102, 51)
wb_listas=openpyxl.load_workbook('Config.xlsx')
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
        btn_logistico.Bind(wx.EVT_BUTTON, self.open_logistica21)
        
        btn_logistico = wx.Button(self.panel, id=wx.ID_ANY, label="Configuracion",size=(-1,-1))
        self.fgs.Add(btn_logistico, pos=(17,6),span=(1,1), flag= wx.ALL, border=0)
        btn_logistico.Bind(wx.EVT_BUTTON, self.configuracion)
        
        
        
        mainSizer= wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(self.fgs,0, flag=wx.ALIGN_CENTER)
        self.panel.SetSizerAndFit(mainSizer)
            
    #-------------Button Functions-----------------#
    def open_nuevo_req11(self, event):
        ww_nuevo_requerimiento11(parent=self.panel).Show()
       

    def open_logistica21(self, event):
        ww_logistica21(parent=self.panel).Show()
        
        
    def configuracion(self, event):
        ww_configuracion(parent=self.panel).Show()

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
        
        #try:
        hist_req_sheet=wb_req['Requerimientos']
        area_req=self.combo_area.GetValue()
        self.fila_vacia = 2
    
        while (hist_req_sheet.cell(row = self.fila_vacia, column = 1).value != None) :
          self.fila_vacia += 1
        
        hist_req_sheet.cell(row=self.fila_vacia, column=col_area_req).value=area_req  
        self.Destroy()

        ww_nuevo_requerimiento12(parent=self.panel).Show()
        pub.sendMessage(event,"panel_listener",message=area_req)
        print(2)
        
        
        #except:
         #   error_msgbox=wx.MessageDialog(None,'Error al guardar el registro en la BD. \nVerifique el el archivo de excel este cerrado y en la ruta correcta.','ERROR',wx.ICON_ERROR)
          #  error_msgbox.ShowModal()
    #-------------Button Functions-----------------# 
        
#############----------------------------------------FRONT END----------------------------------------#############

class ww_nuevo_requerimiento12(wx.Frame):
    
    
    def __init__(self,parent):
        ######----------------------------------------BACK END----------------------------------------#############
        
        pub.subscribe(self.adicionar_transp, "panel_listener")
        print(1)
        
        
        req2_sheet=wb_listas['Requerimientos-12']
        hist_req_sheet=wb_req['Requerimientos']
        
        
        self.lista_encargado=[]
        self.lista_tipo_cont=[]
        self.lista_tipo_transp=[]
        self.lista_descargue=[]
        self.lista_debe_enviarinfo=[]
        self.lista_nro_req=[]
        
        for cell in req2_sheet['A']:
            if cell.value != None:
                self.lista_encargado.append(cell.value)
        for cell in req2_sheet['B']:
            if cell.value != None:
                self.lista_tipo_cont.append(cell.value)
        for cell in req2_sheet['C']:
            if cell.value != None:
                self.lista_tipo_transp.append(cell.value)
        for cell in req2_sheet['E']:
            if cell.value != None:
                self.lista_descargue.append(cell.value)   
        for cell in req2_sheet['F']:
            if cell.value != None:
                self.lista_debe_enviarinfo.append(cell.value)
        
        for cell in hist_req_sheet['A']:
            if cell.value !=None:
                self.lista_nro_req.append(cell.value)
        
        try:
            self.nro_req= int(self.lista_nro_req[-1])+1
        except:
            self.nro_req=1

                
        self.lista_tipo_cont.pop(0)
        self.lista_tipo_transp.pop(0)
        self.lista_descargue.pop(0)
        self.lista_debe_enviarinfo.pop(0)
        
        self.fila_vacia = 2
        
        while (hist_req_sheet.cell(row = self.fila_vacia, column = 1).value != None) :
          self.fila_vacia += 1
 
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
        self.lblrequerimiento_auto =wx.StaticText(self.panel, label=str(self.nro_req))
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
        
        self.combotipotransporte=wx.ComboBox(self.panel,value=self.lista_tipo_transp[0], choices=self.lista_tipo_transp)
        self.combotipocontenedor=wx.ComboBox(self.panel,value=self.lista_tipo_cont[0], choices=self.lista_tipo_cont)
        self.comborequieredescargue=wx.ComboBox(self.panel,value=self.lista_descargue[0], choices=self.lista_descargue)
        self.combodebeinfo=wx.ComboBox(self.panel,value=self.lista_debe_enviarinfo[0], choices=self.lista_debe_enviarinfo)
        
        self.check_si_peaje = wx.CheckBox(self.panel, label= "Si")
        self.check_no_peaje = wx.CheckBox(self.panel, label='No')
        self.check_si_peaje.SetForegroundColour('white')
        self.check_no_peaje.SetForegroundColour('white')
        
        btn_guardar = wx.Button(self.panel, id=wx.ID_ANY, label="Guardar",size=(-1,-1))
        btn_salir = wx.Button(self.panel, id=wx.ID_ANY, label="Salir",size=(-1,-1))
        btn_adicionar_transp = wx.Button(self.panel, id=wx.ID_ANY, label="Adicionar",size=(-1,-1))
        
        self.fgs.Add(self.check_si_peaje, pos=(7,7),span=(1,1), flag= wx.ALL | wx.ALIGN_BOTTOM | wx.ALIGN_CENTER_HORIZONTAL, border=5)
        self.fgs.Add(self.check_no_peaje, pos=(8,7),span=(1,1), flag= wx.LEFT | wx.ALIGN_TOP | wx.ALIGN_CENTER_HORIZONTAL, border=7)
        
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
        self.fgs.Add(self.lblrecargopeaje , pos=(7,6),span=(2,1), flag= wx.ALL| wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
        self.fgs.Add(self.lblinfocliente , pos=(10,4),span=(1,3), flag= wx.ALL| wx.ALIGN_CENTER, border=5)
        self.fgs.Add(self.lblnombreresponsable , pos=(12,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lbltelefono_resp , pos=(12,4),span=(1,1), flag= wx.ALL | wx.ALIGN_RIGHT, border=5)
        self.fgs.Add(self.lblcargo , pos=(12,7),span=(1,1), flag= wx.ALL| wx.ALIGN_RIGHT, border=5)
        self.fgs.Add(self.lblnombresiso , pos=(14,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lbltelefono_siso , pos=(14,4),span=(1,1), flag= wx.ALL |wx.ALIGN_RIGHT, border=5)
        self.fgs.Add(self.lbldebeinfo , pos=(15,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblhorasantes , pos=(16,1),span=(1,1), flag= wx.ALL, border=5)
        
        
        self.check_si_peaje.Bind(wx.EVT_CHECKBOX, self.onCheck_si_peaje)
        self.check_no_peaje.Bind(wx.EVT_CHECKBOX, self.onCheck_no_peaje)
        
        btn_guardar.Bind(wx.EVT_BUTTON, self.guardar_req)
        btn_salir.Bind(wx.EVT_BUTTON, self.salir)
        btn_adicionar_transp.Bind(wx.EVT_BUTTON, self.adicionar_transp)
        
        mainSizer= wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(self.fgs,0, flag=wx.ALIGN_LEFT)
        self.panel.SetSizerAndFit(mainSizer)
        
    def onCheck_si_peaje(self,event):
        if self.check_no_peaje.IsChecked():
            self.check_no_peaje.SetValue(False)
            
    def onCheck_no_peaje(self,event):
        if self.check_si_peaje.IsChecked():
            self.check_si_peaje.SetValue(False)
        
        
    def guardar_req(self,event):
        
        hist_req_sheet=wb_req['Requerimientos']
        req2_sheet=wb_listas['Requerimientos-12']
        
        self.fila_vacia = 1
        
        while (hist_req_sheet.cell(row = self.fila_vacia, column = 1).value != None) :
          self.fila_vacia += 1
        
        self.nro_req=self.fila_vacia-1

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
        
        if self.check_si_peaje.IsChecked():
            check_peaje="Si"
        else:
            check_peaje="No"
         
        self.dic_asosiacion={}
        self.lista_asociacion=[]
        self.lista_tipo_transp2=[]
        
        for cell in req2_sheet['D']:
            if cell != None:
                self.lista_asociacion.append(cell.value)
        
        for cell in req2_sheet['C']:
            if cell != None:
                self.lista_tipo_transp2.append(cell.value)
        
        
        for i in range((len(self.lista_tipo_transp2))):
            self.dic_asosiacion[self.lista_tipo_transp2[i]]=self.lista_asociacion[i]

        hist_req_sheet.cell(row=self.fila_vacia, column=col_requerimiento_auto).value=requerimiento_auto
        hist_req_sheet.cell(row=self.fila_vacia, column=col_area).value=self.dic_asosiacion[tipotransporte]
        hist_req_sheet.cell(row=self.fila_vacia, column=col_fecha_auto).value=fecha_auto
        hist_req_sheet.cell(row=self.fila_vacia, column=col_cotizacion).value=cotizacion
        hist_req_sheet.cell(row=self.fila_vacia, column=col_tipotransporte).value=tipotransporte
        hist_req_sheet.cell(row=self.fila_vacia, column=col_tipocontenedor).value=tipocontenedor
        hist_req_sheet.cell(row=self.fila_vacia, column=col_requieredescargue).value=requieredescargue
        hist_req_sheet.cell(row=self.fila_vacia, column=col_origen).value=origen
        hist_req_sheet.cell(row=self.fila_vacia, column=col_destino).value=destino
        hist_req_sheet.cell(row=self.fila_vacia, column=col_km).value=km
        hist_req_sheet.cell(row=self.fila_vacia, column=col_precio).value=precio
        hist_req_sheet.cell(row=self.fila_vacia, column=col_nombreresponsable).value=nombreresponsable
        hist_req_sheet.cell(row=self.fila_vacia, column=col_telefono_resp).value=telefono_resp
        hist_req_sheet.cell(row=self.fila_vacia, column=col_cargo).value=cargo
        hist_req_sheet.cell(row=self.fila_vacia, column=col_nombresiso).value=nombresiso
        hist_req_sheet.cell(row=self.fila_vacia, column=col_telefono_siso).value=telefono_siso
        hist_req_sheet.cell(row=self.fila_vacia, column=col_debeinfo).value=debeinfo
        hist_req_sheet.cell(row=self.fila_vacia, column=col_horasantes).value=horasantes
        hist_req_sheet.cell(row=self.fila_vacia, column=col_recargopeaje).value=check_peaje
        
        self.txtcotizacion.Value=''
        self.combotipotransporte.Value=self.lista_tipo_transp[0]
        self.combotipocontenedor.Value=self.lista_tipo_cont[0]
        self.comborequieredescargue.Value=self.lista_descargue[0]
        self.txtorigen.Value=''
        self.txtdestino.Value=''
        self.txtkm.Value=''
        self.txtprecio.Value=''
        self.txtnombreresponsable.Value=''
        self.txttelefono_resp.Value=''
        self.txtcargo.Value=''
        self.txtnombresiso.Value=''
        self.txttelefono_siso.Value=''
        self.combodebeinfo.Value=self.lista_debe_enviarinfo[0]
        self.txthorasantes.Value=''
        self.check_no_peaje.SetValue(False)
        self.check_si_peaje.SetValue(False)
        
        try:
            wb_req.save('db_req.xlsx')
            for cell in hist_req_sheet['A']:
                if cell.value !=None:
                    self.lista_nro_req.append(cell.value)
            self.nro_req= int(self.lista_nro_req[-1])+1    
            self.Destroy()
            self.lblrequerimiento_auto.SetLabel(str(self.nro_req))
            
        except:
            error_msgbox=wx.MessageDialog(None,'Error al guardar el registro en la BD. \nVerifique el el archivo de excel este cerrado y en la ruta correcta.','ERROR',wx.ICON_ERROR)
            error_msgbox.ShowModal()


    def salir(self,event):
        self.Destroy()
    
    def adicionar_transp(self,event,message):
        hist_req_sheet=wb_req['Requerimientos']
        req2_sheet=wb_listas['Requerimientos-12']
        print('hola')
        print({message})
        self.fila_vacia = 1
        
        while (hist_req_sheet.cell(row = self.fila_vacia, column = 1).value != None) :
          self.fila_vacia += 1
        
        self.nro_req=self.fila_vacia-1
        
        ultima_area=hist_req_sheet.cell(row=(self.fila_vacia), column=col_area_req).value
        
        
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
        
        if self.check_si_peaje.IsChecked():
            check_peaje="Si"
        else:
            check_peaje="No"
            
        self.dic_asosiacion={}
        self.lista_asociacion=[]
        self.lista_tipo_transp2=[]
        
        for cell in req2_sheet['D']:
            if cell != None:
                self.lista_asociacion.append(cell.value)
        
        for cell in req2_sheet['C']:
            if cell != None:
                self.lista_tipo_transp2.append(cell.value)
        
        
        for i in range((len(self.lista_tipo_transp2))):
            self.dic_asosiacion[self.lista_tipo_transp2[i]]=self.lista_asociacion[i]
        
        hist_req_sheet.cell(row=self.fila_vacia, column=col_requerimiento_auto).value=requerimiento_auto
        hist_req_sheet.cell(row=self.fila_vacia, column=col_area_req).value=ultima_area
        hist_req_sheet.cell(row=self.fila_vacia, column=col_area).value=self.dic_asosiacion[tipotransporte]
        hist_req_sheet.cell(row=self.fila_vacia, column=col_fecha_auto).value=fecha_auto
        hist_req_sheet.cell(row=self.fila_vacia, column=col_cotizacion).value=cotizacion
        hist_req_sheet.cell(row=self.fila_vacia, column=col_tipotransporte).value=tipotransporte
        hist_req_sheet.cell(row=self.fila_vacia, column=col_tipocontenedor).value=tipocontenedor
        hist_req_sheet.cell(row=self.fila_vacia, column=col_requieredescargue).value=requieredescargue
        hist_req_sheet.cell(row=self.fila_vacia, column=col_origen).value=origen
        hist_req_sheet.cell(row=self.fila_vacia, column=col_destino).value=destino
        hist_req_sheet.cell(row=self.fila_vacia, column=col_km).value=km
        hist_req_sheet.cell(row=self.fila_vacia, column=col_precio).value=precio
        hist_req_sheet.cell(row=self.fila_vacia, column=col_nombreresponsable).value=nombreresponsable
        hist_req_sheet.cell(row=self.fila_vacia, column=col_telefono_resp).value=telefono_resp
        hist_req_sheet.cell(row=self.fila_vacia, column=col_cargo).value=cargo
        hist_req_sheet.cell(row=self.fila_vacia, column=col_nombresiso).value=nombresiso
        hist_req_sheet.cell(row=self.fila_vacia, column=col_telefono_siso).value=telefono_siso
        hist_req_sheet.cell(row=self.fila_vacia, column=col_debeinfo).value=debeinfo
        hist_req_sheet.cell(row=self.fila_vacia, column=col_horasantes).value=horasantes
        hist_req_sheet.cell(row=self.fila_vacia, column=col_recargopeaje).value=check_peaje
        
        self.combotipotransporte.Value=self.lista_tipo_transp[0]
        self.combotipocontenedor.Value=self.lista_tipo_cont[0]
        self.txtprecio.Value=''
        
        try:
            wb_req.save('db_req.xlsx')
            for cell in hist_req_sheet['A']:
                if cell.value !=None:
                    self.lista_nro_req.append(cell.value)
            self.nro_req= int(self.lista_nro_req[-1])+1        
            self.lblrequerimiento_auto.SetLabel(str(self.nro_req))
        except:
            error_msgbox=wx.MessageDialog(None,'Error al guardar el registro en la BD. \nVerifique el el archivo de excel este cerrado y en la ruta correcta.','ERROR',wx.ICON_ERROR)
            error_msgbox.ShowModal()        

class ww_logistica21(wx.Frame):
    
    def __init__(self,parent):
        ######----------------------------------------BACK END----------------------------------------#############        

        ######----------------------------------------BACK END----------------------------------------#############       
        
        ######----------------------------------------FRONT END----------------------------------------#############
        
        wx.Frame.__init__(self, None, wx.ID_ANY, "Contenedores de Antioquia - Centro Logistico", size=(270, 250))  
        
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
       
        self.lbltitle =wx.StaticText(self.panel, label='Ingrese Numero de Requerimiento\n a Gestionar:')
        self.lbltitle.SetFont(title_font)
        self.lbltitle.SetBackgroundColour(principal_color)
        self.lbltitle.SetForegroundColour('white')
        self.fgs.Add(self.lbltitle,pos=(2,1),span=(1,3), flag=wx.ALL | wx.ALIGN_CENTER, border=5)

        self.txtreq = wx.TextCtrl(self.panel)
        self.fgs.Add(self.txtreq , pos=(3,1),span=(1,3), flag= wx.ALL | wx.EXPAND, border=5)
        
        btn_aceptar = wx.Button(self.panel, id=wx.ID_ANY, label="Aceptar",size=(-1,-1))
        self.fgs.Add(btn_aceptar, pos=(6,1),span=(1,3), flag= wx.ALL | wx.ALIGN_CENTER, border=0)
        btn_aceptar.Bind(wx.EVT_BUTTON, self.open_logistica22)

        mainSizer= wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(self.fgs,0, flag=wx.ALIGN_LEFT)
        self.panel.SetSizerAndFit(mainSizer)
        
    #-------------Button Functions-----------------#            
    def open_logistica22(self, event):
        
        hist_req_sheet=wb_req['Requerimientos']
        self.lista_nro_req=[]
        
        for cell in hist_req_sheet['A']:
            if cell.value !=None:
                self.lista_nro_req.append(cell.value)
        global req_selec
        req_selec=self.txtreq.GetValue()
                
        if req_selec in self.lista_nro_req:
            ww_nuevo_requerimiento22(parent=self.panel).Show() 
            self.Destroy()
        else:
            error_msgbox=wx.MessageDialog(None,'Error: Numero de Requerimiento No Encontrado','ERROR',wx.ICON_ERROR)
            error_msgbox.ShowModal()    
           
class ww_nuevo_requerimiento22(wx.Frame):
    def __init__(self,parent):    
          
        self.hist_req_sheet=wb_req['Requerimientos']
        global req_selec
        
        self.lista_requerimientos=[]
        
        for cell in self.hist_req_sheet['A']:
            if cell.value != None:
                self.lista_requerimientos.append(cell.value)

        self.nro_fila_req=int(self.lista_requerimientos.index(req_selec))+1
        
        self.lista_valores_fila=[]
        for cell in self.hist_req_sheet[self.nro_fila_req]:
            self.lista_valores_fila.append(cell.value)
        
        #----------Front------------#
        wx.Frame.__init__(self, None, wx.ID_ANY, "Contenedores de Antioquia - Centro Logistico", size=(950, 600))
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
    
        self.lblrequerimiento=wx.StaticText(self.panel, label='Requerimiento N°')
        self.lblfecha=wx.StaticText(self.panel, label='Fecha')
        self.lblareaencargada=wx.StaticText(self.panel, label='Area Encargada')
        self.lblcotizacion=wx.StaticText(self.panel, label='Cotizacion N°')
        self.lbltipotransp=wx.StaticText(self.panel, label='Tipo de Transporte')
        self.lbltipocont=wx.StaticText(self.panel, label='Tipo de Contenedor')
        self.lbldescargue=wx.StaticText(self.panel, label='Requiere Descargue')
        self.lblorigen=wx.StaticText(self.panel, label='Origen')
        self.lbldestino=wx.StaticText(self.panel, label='Destino')
        self.lblkm=wx.StaticText(self.panel, label='Km')
        self.lblprecio=wx.StaticText(self.panel, label='Precio')
        self.lblrecargopeaje=wx.StaticText(self.panel, label='Recargo Peaje')
        self.lblnombreresp=wx.StaticText(self.panel, label='Nombre\nResponsable')
        self.lbltelresp=wx.StaticText(self.panel, label='Telefono Resp.')
        self.lblcargoresp=wx.StaticText(self.panel, label='Cargo')
        self.lblnombresiso=wx.StaticText(self.panel, label='Nombre SISO')
        self.lbltelesiso=wx.StaticText(self.panel, label='Telefono SISO')
        self.lbldebeinfo=wx.StaticText(self.panel, label='Debe Enviarse\nInformacion')
        self.lblhorasantes=wx.StaticText(self.panel, label='N° de Horas Antes')
        self.lblinfologistica=wx.StaticText(self.panel, label='Info Logistica')
        self.lblinfocliente=wx.StaticText(self.panel, label='Info Cliente')
        self.lblfechaentrega=wx.StaticText(self.panel, label='Fecha de Entrega')
        self.lbldireccion=wx.StaticText(self.panel, label='Direccion Exacta')
        self.lblreferenciacont=wx.StaticText(self.panel, label='Ref.Contenedor')
        self.lblnombreconduc=wx.StaticText(self.panel, label='Nombre Conductor')
        self.lblcedula=wx.StaticText(self.panel, label='Cedula')
        self.lbltelefonoconduc=wx.StaticText(self.panel, label='Telefono')
        self.lblplaca=wx.StaticText(self.panel, label='Placa')
        self.lbladiciones=wx.StaticText(self.panel, label='Adiciones Entrega')
        self.lblpreguntahoras=wx.StaticText(self.panel, label='Documentacion Enviada x Horas Antes?')
        self.lblpreguntadoc=wx.StaticText(self.panel, label='Documentacion Completa?')
        self.lblrptarequerimiento=wx.StaticText(self.panel, label=self.lista_valores_fila[col_requerimiento_auto-1])
        self.lblrptafecha=wx.StaticText(self.panel, label=self.lista_valores_fila[col_fecha_auto-1])
        self.lblrptaareaencargada=wx.StaticText(self.panel, label=self.lista_valores_fila[col_area-1])
        self.lblrptacotizacion=wx.StaticText(self.panel, label=self.lista_valores_fila[col_cotizacion-1])
        self.lblrptatipotransp=wx.StaticText(self.panel, label=self.lista_valores_fila[col_tipotransporte-1])
        self.lblrptatipocont=wx.StaticText(self.panel, label=self.lista_valores_fila[col_tipocontenedor-1])
        self.lblrptadescargue=wx.StaticText(self.panel, label=self.lista_valores_fila[col_requieredescargue-1])
        self.lblrptaorigen=wx.StaticText(self.panel, label=self.lista_valores_fila[col_origen-1])
        self.lblrptadestino=wx.StaticText(self.panel, label=self.lista_valores_fila[col_destino-1])
        self.lblrptakm=wx.StaticText(self.panel, label=self.lista_valores_fila[col_km-1])
        self.lblrptaprecio=wx.StaticText(self.panel, label=self.lista_valores_fila[col_precio-1])
        self.lblrptarecargopeaje=wx.StaticText(self.panel, label=self.lista_valores_fila[col_recargopeaje-1])
        self.lblrptanombreresp=wx.StaticText(self.panel, label=self.lista_valores_fila[col_nombreresponsable-1])
        self.lblrptatelresp=wx.StaticText(self.panel, label=self.lista_valores_fila[col_telefono_resp-1])
        self.lblrptacargoresp=wx.StaticText(self.panel, label=self.lista_valores_fila[col_cargo-1])
        self.lblrptanombresiso=wx.StaticText(self.panel, label=self.lista_valores_fila[col_nombresiso-1])
        self.lblrptatelesiso=wx.StaticText(self.panel, label=self.lista_valores_fila[col_telefono_siso-1])
        self.lblrptadebeinfo=wx.StaticText(self.panel, label=self.lista_valores_fila[col_debeinfo-1])
        self.lblrptahorasantes=wx.StaticText(self.panel, label=str(self.lista_valores_fila[col_horasantes-1]))
        self.txtfechaentrega=wx.TextCtrl(self.panel)
        self.txtdireccion=wx.TextCtrl(self.panel)
        self.txtreferenciacont=wx.TextCtrl(self.panel)
        self.txtnombreconduc=wx.TextCtrl(self.panel)
        self.txtcedula=wx.TextCtrl(self.panel)
        self.txttelefonoconduc=wx.TextCtrl(self.panel)
        self.txtplaca=wx.TextCtrl(self.panel)
        self.txtadiciones=wx.TextCtrl(self.panel,style = wx.TE_MULTILINE)
        self.checkpreguntahoras_si=wx.CheckBox(self.panel, label= "Si")
        self.checkpreguntahoras_no=wx.CheckBox(self.panel, label= "No")
        self.checkpreguntadoc_si=wx.CheckBox(self.panel, label= "Si")
        self.checkpreguntadoc_no=wx.CheckBox(self.panel, label= "No")
    
        self.lblrequerimiento.SetBackgroundColour(principal_color)
        self.lblfecha.SetBackgroundColour(principal_color)
        self.lblareaencargada.SetBackgroundColour(principal_color)
        self.lblcotizacion.SetBackgroundColour(principal_color)
        self.lbltipotransp.SetBackgroundColour(principal_color)
        self.lbltipocont.SetBackgroundColour(principal_color)
        self.lbldescargue.SetBackgroundColour(principal_color)
        self.lblorigen.SetBackgroundColour(principal_color)
        self.lbldestino.SetBackgroundColour(principal_color)
        self.lblkm.SetBackgroundColour(principal_color)
        self.lblprecio.SetBackgroundColour(principal_color)
        self.lblrecargopeaje.SetBackgroundColour(principal_color)
        self.lblnombreresp.SetBackgroundColour(principal_color)
        self.lbltelresp.SetBackgroundColour(principal_color)
        self.lblcargoresp.SetBackgroundColour(principal_color)
        self.lblnombresiso.SetBackgroundColour(principal_color)
        self.lbltelesiso.SetBackgroundColour(principal_color)
        self.lbldebeinfo.SetBackgroundColour(principal_color)
        self.lblhorasantes.SetBackgroundColour(principal_color)
        self.lblinfologistica.SetBackgroundColour(principal_color)
        self.lblinfocliente.SetBackgroundColour(principal_color)
        self.lblfechaentrega.SetBackgroundColour(principal_color)
        self.lbldireccion.SetBackgroundColour(principal_color)
        self.lblreferenciacont.SetBackgroundColour(principal_color)
        self.lblnombreconduc.SetBackgroundColour(principal_color)
        self.lblcedula.SetBackgroundColour(principal_color)
        self.lbltelefonoconduc.SetBackgroundColour(principal_color)
        self.lblplaca.SetBackgroundColour(principal_color)
        self.lbladiciones.SetBackgroundColour(principal_color)
        self.lblpreguntahoras.SetBackgroundColour(principal_color)
        self.lblpreguntadoc.SetBackgroundColour(principal_color)
        self.lblrptarequerimiento.SetBackgroundColour(principal_color)
        self.lblrptafecha.SetBackgroundColour(principal_color)
        self.lblrptaareaencargada.SetBackgroundColour(principal_color)
        self.lblrptacotizacion.SetBackgroundColour(principal_color)
        self.lblrptatipotransp.SetBackgroundColour(principal_color)
        self.lblrptatipocont.SetBackgroundColour(principal_color)
        self.lblrptadescargue.SetBackgroundColour(principal_color)
        self.lblrptaorigen.SetBackgroundColour(principal_color)
        self.lblrptadestino.SetBackgroundColour(principal_color)
        self.lblrptakm.SetBackgroundColour(principal_color)
        self.lblrptaprecio.SetBackgroundColour(principal_color)
        self.lblrptarecargopeaje.SetBackgroundColour(principal_color)
        self.lblrptanombreresp.SetBackgroundColour(principal_color)
        self.lblrptatelresp.SetBackgroundColour(principal_color)
        self.lblrptacargoresp.SetBackgroundColour(principal_color)
        self.lblrptanombresiso.SetBackgroundColour(principal_color)
        self.lblrptatelesiso.SetBackgroundColour(principal_color)
        self.lblrptadebeinfo.SetBackgroundColour(principal_color)
        self.lblrptahorasantes.SetBackgroundColour(principal_color)
        self.checkpreguntahoras_si.SetBackgroundColour(principal_color)
        self.checkpreguntahoras_no.SetBackgroundColour(principal_color)
        self.checkpreguntadoc_si.SetBackgroundColour(principal_color)
        self.checkpreguntadoc_no.SetBackgroundColour(principal_color)

        self.lblrequerimiento.SetForegroundColour('white')
        self.lblfecha.SetForegroundColour('white')
        self.lblareaencargada.SetForegroundColour('white')
        self.lblcotizacion.SetForegroundColour('white')
        self.lbltipotransp.SetForegroundColour('white')
        self.lbltipocont.SetForegroundColour('white')
        self.lbldescargue.SetForegroundColour('white')
        self.lblorigen.SetForegroundColour('white')
        self.lbldestino.SetForegroundColour('white')
        self.lblkm.SetForegroundColour('white')
        self.lblprecio.SetForegroundColour('white')
        self.lblrecargopeaje.SetForegroundColour('white')
        self.lblnombreresp.SetForegroundColour('white')
        self.lbltelresp.SetForegroundColour('white')
        self.lblcargoresp.SetForegroundColour('white')
        self.lblnombresiso.SetForegroundColour('white')
        self.lbltelesiso.SetForegroundColour('white')
        self.lbldebeinfo.SetForegroundColour('white')
        self.lblhorasantes.SetForegroundColour('white')
        self.lblinfologistica.SetForegroundColour('white')
        self.lblinfocliente.SetForegroundColour('white')
        self.lblfechaentrega.SetForegroundColour('white')
        self.lbldireccion.SetForegroundColour('white')
        self.lblreferenciacont.SetForegroundColour('white')
        self.lblnombreconduc.SetForegroundColour('white')
        self.lblcedula.SetForegroundColour('white')
        self.lbltelefonoconduc.SetForegroundColour('white')
        self.lblplaca.SetForegroundColour('white')
        self.lbladiciones.SetForegroundColour('white')
        self.lblpreguntahoras.SetForegroundColour('white')
        self.lblpreguntadoc.SetForegroundColour('white')
        self.lblrptarequerimiento.SetForegroundColour('white')
        self.lblrptafecha.SetForegroundColour('white')
        self.lblrptaareaencargada.SetForegroundColour('white')
        self.lblrptacotizacion.SetForegroundColour('white')
        self.lblrptatipotransp.SetForegroundColour('white')
        self.lblrptatipocont.SetForegroundColour('white')
        self.lblrptadescargue.SetForegroundColour('white')
        self.lblrptaorigen.SetForegroundColour('white')
        self.lblrptadestino.SetForegroundColour('white')
        self.lblrptakm.SetForegroundColour('white')
        self.lblrptaprecio.SetForegroundColour('white')
        self.lblrptarecargopeaje.SetForegroundColour('white')
        self.lblrptanombreresp.SetForegroundColour('white')
        self.lblrptatelresp.SetForegroundColour('white')
        self.lblrptacargoresp.SetForegroundColour('white')
        self.lblrptanombresiso.SetForegroundColour('white')
        self.lblrptatelesiso.SetForegroundColour('white')
        self.lblrptadebeinfo.SetForegroundColour('white')
        self.lblrptahorasantes.SetForegroundColour('white')
        self.checkpreguntahoras_si.SetForegroundColour('white')
        self.checkpreguntahoras_no.SetForegroundColour('white')
        self.checkpreguntadoc_si.SetForegroundColour('white')
        self.checkpreguntadoc_no.SetForegroundColour('white')
        
        btn_imprimir = wx.Button(self.panel, id=wx.ID_ANY, label="Imprimir Remision",size=(-1,-1))
        btn_finalizar = wx.Button(self.panel, id=wx.ID_ANY, label="Finalizar",size=(-1,-1))

        self.fgs.Add(self.lblrequerimiento, pos=(1,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblfecha, pos=(2,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblareaencargada, pos=(1,7),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblcotizacion, pos=(5,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lbltipotransp, pos=(6,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lbltipocont, pos=(7,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lbldescargue, pos=(8,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblorigen, pos=(5,3),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lbldestino, pos=(6,3),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblkm, pos=(7,3),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblprecio, pos=(8,3),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrecargopeaje, pos=(9,3),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblnombreresp, pos=(5,5),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lbltelresp, pos=(6,5),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblcargoresp, pos=(7,5),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblnombresiso, pos=(8,5),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lbltelesiso, pos=(9,5),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lbldebeinfo, pos=(5,7),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblhorasantes, pos=(6,7),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblinfologistica, pos=(4,1),span=(1,4), flag= wx.ALL | wx.ALIGN_CENTER, border=5)
        self.fgs.Add(self.lblinfocliente, pos=(4,5),span=(1,4), flag= wx.ALL| wx.ALIGN_CENTER, border=5)
        self.fgs.Add(self.lblfechaentrega, pos=(11,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lbldireccion, pos=(11,3),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblreferenciacont, pos=(11,7),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblnombreconduc, pos=(12,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblcedula, pos=(12,2),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lbltelefonoconduc, pos=(12,3),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblplaca, pos=(12,4),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lbladiciones, pos=(12,6),span=(1,3), flag= wx.ALL|wx.ALIGN_CENTER, border=5)
        self.fgs.Add(self.lblpreguntahoras, pos=(15,1),span=(2,2), flag= wx.ALL |wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL , border=5)
        self.fgs.Add(self.lblpreguntadoc, pos=(18,1),span=(2,2), flag= wx.ALL |wx.ALIGN_RIGHT | wx.ALIGN_CENTER_VERTICAL, border=5)
        self.fgs.Add(self.lblrptarequerimiento, pos=(1,2),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrptafecha, pos=(2,2),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrptaareaencargada, pos=(1,8),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrptacotizacion, pos=(5,2),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrptatipotransp, pos=(6,2),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrptatipocont, pos=(7,2),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrptadescargue, pos=(8,2),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrptaorigen, pos=(5,4),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrptadestino, pos=(6,4),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrptakm, pos=(7,4),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrptaprecio, pos=(8,4),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrptarecargopeaje, pos=(9,4),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrptanombreresp, pos=(5,6),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrptatelresp, pos=(6,6),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrptacargoresp, pos=(7,6),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrptanombresiso, pos=(8,6),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrptatelesiso, pos=(9,6),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrptadebeinfo, pos=(5,8),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.lblrptahorasantes, pos=(6,8),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.txtfechaentrega, pos=(11,2),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.txtdireccion, pos=(11,4),span=(1,3), flag= wx.ALL, border=5)
        self.fgs.Add(self.txtreferenciacont, pos=(11,8),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.txtnombreconduc, pos=(13,1),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.txtcedula, pos=(13,2),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.txttelefonoconduc, pos=(13,3),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.txtplaca, pos=(13,4),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.txtadiciones, pos=(13,6),span=(4,3), flag= wx.ALL| wx.EXPAND, border=5)
        self.fgs.Add(self.checkpreguntahoras_si, pos=(15,3),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.checkpreguntahoras_no, pos=(16,3),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.checkpreguntadoc_si, pos=(18,3),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(self.checkpreguntadoc_no, pos=(19,3),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(btn_finalizar, pos=(19,8),span=(1,1), flag= wx.ALL, border=5)
        self.fgs.Add(btn_imprimir, pos=(19,7),span=(1,1), flag= wx.ALL, border=5)
        
        self.checkpreguntahoras_si.Bind(wx.EVT_CHECKBOX, self.onCheckhoras_si)
        self.checkpreguntahoras_no.Bind(wx.EVT_CHECKBOX, self.onCheckhoras_no)
        
        self.checkpreguntadoc_si.Bind(wx.EVT_CHECKBOX, self.onCheckdoc_si)
        self.checkpreguntadoc_no.Bind(wx.EVT_CHECKBOX, self.onCheckdoc_no)
        
        
        
        btn_finalizar.Bind(wx.EVT_BUTTON, self.finalizar)
        btn_imprimir.Bind(wx.EVT_BUTTON, self.imprimir)
       
        mainSizer= wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(self.fgs,0, flag=wx.ALIGN_LEFT)
        self.panel.SetSizerAndFit(mainSizer)
        

  
    def onCheckhoras_si(self,event):
        if self.checkpreguntahoras_no.IsChecked():
            self.checkpreguntahoras_no.SetValue(False)
            
    def onCheckhoras_no(self,event):
        if self.checkpreguntahoras_si.IsChecked():
            self.checkpreguntahoras_si.SetValue(False)
            
    def onCheckdoc_si(self,event):
        if self.checkpreguntadoc_no.IsChecked():
            self.checkpreguntadoc_no.SetValue(False)
            
    def onCheckdoc_no(self,event):
        if self.checkpreguntadoc_si.IsChecked():
            self.checkpreguntadoc_si.SetValue(False)
            
            
        
    def finalizar(self,event):
        
        fechaentrega=self.txtfechaentrega.GetValue()
        direccion=self.txtdireccion.GetValue()
        referenciacont=self.txtreferenciacont.GetValue()
        nombreconduc=self.txtnombreconduc.GetValue()
        cedula=self.txtcedula.GetValue()
        telefonoconduc=self.txttelefonoconduc.GetValue()
        placa=self.txtplaca.GetValue()
        adiciones=self.txtadiciones.GetValue()
        
        if self.checkpreguntahoras_si.IsChecked():
            check_horas="Si"
        else:
            check_peaje="No"
        
        if self.checkpreguntadoc_si.IsChecked():
            check_doc="Si"
        else:
            check_doc="No"

        self.hist_req_sheet.cell(row=self.nro_fila_req, column=col_fechaentrega).value=fechaentrega
        self.hist_req_sheet.cell(row=self.nro_fila_req, column=col_direccion).value=direccion
        self.hist_req_sheet.cell(row=self.nro_fila_req, column=col_referenciacont).value=referenciacont
        self.hist_req_sheet.cell(row=self.nro_fila_req, column=col_nombreconduc).value=nombreconduc
        self.hist_req_sheet.cell(row=self.nro_fila_req, column=col_cedula).value=cedula
        self.hist_req_sheet.cell(row=self.nro_fila_req, column=col_telefonoconduc).value=telefonoconduc
        self.hist_req_sheet.cell(row=self.nro_fila_req, column=col_placa).value=placa
        self.hist_req_sheet.cell(row=self.nro_fila_req, column=col_adiciones).value=adiciones
        self.hist_req_sheet.cell(row=self.nro_fila_req, column=col_preguntahoras).value=check_horas
        self.hist_req_sheet.cell(row=self.nro_fila_req, column=col_preguntadoc).value=check_doc

        try:
            wb_req.save('db_req.xlsx')
            sgto_msgbox=wx.MessageDialog(None,'Recuerde Hacer el Seguimiento','Recuerde!',wx.ICON_WARNING)
            sgto_msgbox.ShowModal()
            self.Destroy()
        except:
            error_msgbox=wx.MessageDialog(None,'Error al guardar el registro en la BD. \nVerifique el el archivo de excel este cerrado y en la ruta correcta.','ERROR',wx.ICON_ERROR)
            error_msgbox.ShowModal()
        
        
    def imprimir(self, event):
        pass
    
class ww_configuracion(wx.Frame):   
    
    def __init__(self,parent):
   
        
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
       
        self.lbltitle =wx.StaticText(self.panel, label='Ingrese Contraseña')
        self.lbltitle.SetFont(title_font)
        self.lbltitle.SetBackgroundColour(principal_color)
        self.lbltitle.SetForegroundColour('white')
        self.fgs.Add(self.lbltitle,pos=(2,1),span=(1,3), flag=wx.ALL | wx.ALIGN_CENTER, border=5)

        self.txtpass = wx.TextCtrl(self.panel, style= wx.TE_PASSWORD)
        self.fgs.Add(self.txtpass , pos=(3,1),span=(1,3), flag= wx.ALL | wx.EXPAND, border=5)
        
        btn_aceptar = wx.Button(self.panel, id=wx.ID_ANY, label="Aceptar",size=(-1,-1))
        self.fgs.Add(btn_aceptar, pos=(6,1),span=(1,3), flag= wx.ALL | wx.ALIGN_CENTER, border=0)
        btn_aceptar.Bind(wx.EVT_BUTTON, self.onBtn_aceptar)

        mainSizer= wx.BoxSizer(wx.VERTICAL)
        mainSizer.Add(self.fgs,0, flag=wx.ALIGN_CENTER)
        self.panel.SetSizerAndFit(mainSizer)
        
    #-------------Button Functions-----------------#            
    def onBtn_aceptar(self, event):
        self.config_sheet=wb_listas['Config']
        
        self.Destroy()
        #ww_nuevo_requerimiento12(parent=self.panel).Show()

    #-------------Button Functions-----------------# 
        
#############----------------------------------------FRONT END----------------------------------------#############
        

class MyApp(wx.App):
    def OnInit(self):
        self.frame= MyFrame()
        self.frame.Show()
        return True       
 
# Run the program     
app=MyApp()
app.MainLoop()
del app
            
            