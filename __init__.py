#! /usr/bin/env python
# ####! -*- coding: utf-8 -*-

import gtk
import os
import xlwt
import threading
from datetime import datetime, date
from serial import Serial
import subprocess
import time


class ThreadGtk(threading.Thread):
    """Hilo para actualizar etiqueta"""
    def __init__(self, label):
        super(ThreadGtk, self).__init__()
        self.label = label
        self.quit = False
        self.text = ''
        self.event = threading.Event()
        self.start()

    def run(self):
        while not self.quit:
            self.event.wait()
            self.label.push(0, self.text)
            self.event.clear()


class Ventana(gtk.Window):

    def __init__(self):
        super(Ventana, self).__init__()
        self.set_title('MIT-04 v1.2')
        self.set_position(gtk.WIN_POS_CENTER_ALWAYS)
        path = os.path.join('images', 'logo.png')
        icon48 = gtk.gdk.pixbuf_new_from_file(path)
        self.set_icon_list(icon48)
        self.letras = [
            ' ', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I',
            'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S',
            'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
        self.metodos = ['', 'SECUENCIAL', 'COMPARATIVO']
        # Fondo
        path = os.path.join('images', 'fondo.jpg')
        pixbuf = gtk.gdk.pixbuf_new_from_file(path)
        pixmap, mask = pixbuf.render_pixmap_and_mask()
        width, height = pixmap.get_size()
        self.set_size_request(width, height + 0)
        self.set_resizable(False)
        del pixbuf
        self.set_app_paintable(True)
        self.realize()
        self.window.set_back_pixmap(pixmap, False)
        del pixmap
        main_vbox = gtk.VBox(False, 0)
        self.add(main_vbox)
        hbox = gtk.HBox(False, 0)
        main_vbox.pack_start(hbox, True, True, 0)
        self.label = gtk.Statusbar()
        main_vbox.pack_start(self.label, False, False, 0)
        vbox = gtk.VBox(True, 0)
        hbox.pack_end(vbox, False, False, 5)
        self.button = Button('escanear.png', 'Escanear')
        vbox.pack_start(self.button, False, False, 0)
        self.hora = Button('hora.png', 'Act. Hora')
        vbox.pack_start(self.hora, True, False, 0)
        self.borrar_but = Button('borrar.png', 'Borrar Memoria')
        vbox.pack_start(self.borrar_but, True, False, 0)
        self.button.connect('clicked', self.conectar)
        self.hora.connect('clicked', self.actualizar_hora)
        self.borrar_but.connect('clicked', self.borrar)
        borders = xlwt.Borders()
        borders.left = xlwt.Borders.MEDIUM
        borders.right = xlwt.Borders.MEDIUM
        borders.top = xlwt.Borders.MEDIUM
        borders.bottom = xlwt.Borders.MEDIUM
#        pattern = xlwt.Pattern()
#        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
#        pattern.pattern_fore_colour = 31
        alignment = xlwt.Alignment()
        alignment.horz = xlwt.Alignment.HORZ_CENTER
        alignment.vert = xlwt.Alignment.VERT_CENTER
        font = xlwt.Font()
        font.name = 'Arial'
        font.bold = True
        self.estilo_bold = xlwt.XFStyle()
        self.estilo_bold.borders = borders
#        self.estilo_bold.pattern = pattern
        self.estilo_bold.alignment = alignment
        self.estilo_bold.font = font

        borders = xlwt.Borders()
        borders.left = xlwt.Borders.THIN
        borders.right = xlwt.Borders.THIN
        borders.top = xlwt.Borders.THIN
        borders.bottom = xlwt.Borders.THIN
        self.estilo = xlwt.XFStyle()
        self.estilo.borders = borders
#        self.estilo.pattern = pattern
        self.estilo.alignment = alignment
        self.estilo_fecha = xlwt.XFStyle()
        self.estilo_fecha.num_format_str = 'DD-MM-YY'
        self.estilo_fecha.borders = borders
        self.estilo_fecha.alignment = alignment
        self.estilo_hora = xlwt.XFStyle()
        self.estilo_hora.num_format_str = 'HH:MM:SS'
        self.estilo_hora.borders = borders
        self.estilo_hora.alignment = alignment
        self.show_all()
        self.hilogtk = ThreadGtk(self.label)
        self.escanear = True
        hilo = threading.Thread(target=self.escaneando)
        hilo.start()
        self.connect('destroy', self.quit)

    def label_update(self, texto):
        self.hilogtk.text = texto
        self.hilogtk.event.set()

    def desconectar(self):
        self.escanear = True
        self.button.update('escanear.png', 'Escanear')
        self.hora.set_sensitive(False)
        self.borrar_but.set_sensitive(False)
        try:
            self.puerto.close()
        except:
            self.label_update('No se pudo desconectar correctamente')

    def actualizar_hora(self, *args):
        com = 'AT+MD?' + chr(0x0D) + chr(0x0A)
        self.puerto = Serial(port=self.puerto_nombre, baudrate=9600, timeout=1)
        self.puerto.write(com)
        self.label_update('Comunicándose con MIT04')
        try:
            respuesta = self.puerto.read(1101)
        except:
            self.label_update('Nada recibido')
            self.desconectar()
            return
        if not 'MIT03' in respuesta:
            self.label_update('Puerto equivocado')
            self.desconectar()
            return
        now = datetime.now()
        wd = now.isoweekday()
        if wd == 7:
            wd = 0
        pwd = self.tonible(wd)
        hora = self.tonible(now.hour)
        minuto = self.tonible(now.minute)
        segundo = self.tonible(now.second)
        dia = self.tonible(now.day)
        mes = self.tonible(now.month)
        ano = self.tonible(now.year - 2000)
        ascii = hora + minuto + segundo + dia + mes + ano + pwd
        com = 'AT+SETT=' + ascii + chr(0x0D) + chr(0x0A)
        self.puerto.write(com)
        try:
            respuesta = self.puerto.read(1101)
            if 'OK' in respuesta:
                self.label_update('Hora actualizada')
            else:
                self.label_update('Error de conexión')
                self.desconectar()
                return
        except:
            self.label_update('Error de conexión')
            self.desconectar()
            return
        self.puerto.close()

    def conectar(self, *args):
        print 'conectar'
        if self.escanear:
            self.label_update('Escaneando puertos')
            self.escaneando()
            return
        com = 'AT+MD?' + chr(0x0D) + chr(0x0A)
        self.puerto = Serial(port=self.puerto_nombre, baudrate=9600, timeout=1)
        self.puerto.write(com)
        self.label_update('Comunicándose con MIT04')
        try:
            respuesta = self.puerto.read(128)
        except:
            self.label_update('Nada recibido')
            self.desconectar()
            return
        if not 'MIT03' in respuesta:
            self.label_update('Pérdida de Conexión')
            self.desconectar()
            return
        com = 'AT+DWLD?' + chr(0x0D) + chr(0x0A)
        self.puerto.write(com)
        respuesta = ''
        try:
            l = 1
            while l != 0 and len(respuesta) < 32 * 20 * 100:
                parte = self.puerto.read(32 * 20)
                l = len(parte)
                respuesta += parte
                print 'respuesta', len(respuesta), len(respuesta) < 32 * 20 * 100
        except:
            self.label_update('Error de conexión')
            self.desconectar()
            return
        if respuesta == 'OK':
            self.label_update('Registro Vacío')
            self.puerto.close()
            return
        self.puerto.close()

#    def prueba_xls(self, respuesta):
        csv = self.decodificar(respuesta)
        print csv
        self.label_update('Generando Archivo')
        wb = xlwt.Workbook()
        ws = wb.add_sheet('MUESTREO')
        anchos = [1.25, 1.8, 1.8, 1.8, 2.5, 2.5, 2.5, 2.5, 2.5, 3]
        for i, w in enumerate(anchos):
            ws.col(i).width = 1312 * w
        ws.write_merge(0, 1, 0, 9, 'TABLA DE MEDICIONES',
            self.estilo_bold)
        headers = ['Num', 'Fecha', 'Hora', 'Lote', 'Tiempo1 (us)',
            'Tiempo2 (us)', 'Tiempo3 (us)', 'Tiempo4 (us)', 'Tiempo5 (us)', 'Metodo']
        for i, h in enumerate(headers):
            ws.write(3, i, h, self.estilo_bold)
        x = 4
        for f in csv:
            ws.write(x, 0, x - 3, self.estilo)
            ws.write(x, 1, f[1].date(), self.estilo_fecha)
            ws.write(x, 2, f[1].time(), self.estilo_hora)
            ws.write(x, 3, f[0], self.estilo)
            ws.write(x, 9, f[8], self.estilo)
            for i in range(f[7]):
                ws.write(x, i + 4, f[i + 2], self.estilo)
            x += 1
        self.label_update('Archivo generado')
        dia = date.today()
        if os.name == 'nt':
            path = 'outs\MIT04-%s.xls' % dia
            i = 0
            while True:
                try:
                    archivo = open(path, 'rb')
                except:
                    wb.save(path)
                    break
                else:
                    i += 1
                    path = 'outs\MIT03-%s(%d).xls' % (dia, i)
            os.system('start %s' % path)
            self.label_update('Guardado en %s' % path)
        else:
            path = 'outs/MIT04-%s.xls' % dia
            i = 0
            while True:
                try:
                    archivo = open(path, 'rb')
                except:
                    wb.save(path)
                    break
                else:
                    i += 1
                    path = 'outs/MIT04-%s(%d).xls' % (dia, i)
            self.label_update('Guardado en %s' % path)
            os.system('gnome-open %s' % path)
        self.borrar()

    def borrar(self, *args):
        dialogo = Dialogo(self, '¿Borrar Datos?', 'warning.png',
            '¿Desea borrar todos los datos del MIT?')
        if dialogo.iniciar():
            com = 'AT+MD?' + chr(0x0D) + chr(0x0A)
            self.puerto = Serial(port=self.puerto_nombre, baudrate=9600,
                timeout=1)
            self.puerto.write(com)
            self.label_update('Comunicándose con MIT04')
            try:
                respuesta = self.puerto.read(1101)
            except:
                self.label_update('Nada recibido')
                self.desconectar()
                return
            if not 'MIT03' in respuesta:
                self.label_update('Pérdida de Conexión')
                self.desconectar()
                return
            com = 'AT+CLR' + chr(0x0D) + chr(0x0A)
            self.puerto.write(com)
            Alerta(self, 'Operación completada', 'vacio.png',
                'Se borraron todos los registros correctamente')
        dialogo.cerrar()
        self.puerto.close()

    def get_bytes(self, bytes):
        suma = 0
        for i in range(bytes):
            ascii = self.palabra[0:1]
            self.palabra = self.palabra[1:]
            suma = suma * (2 ** 8) + ord(ascii)
        print 'nible', suma
        return suma

    def toint(self, n):
        l = n % 16
        h = n / 16
        return h * 10 + l

    def tonible(self, i):
        l = i % 10
        h = i / 10
        return chr(h * 16 + l + 128)

    def decodificar(self, ascii):
        csv = []
        n = 0
        longitud = len(ascii)
        print 'longitud', longitud
        while True:
            inicio = n * 32
            fin = inicio + 32
            if longitud < fin:
                break
            self.palabra = ascii[inicio:fin]
            try:
                # 5 bytes para el lote
                lote = self.get_bytes(2)
                letras = ''
                for i in range(3):
                    letras = self.letras[lote % 27] + letras
                    lote = lote / 27
                numeros = self.get_bytes(2)
                lote = letras + str(numeros)
                hora = self.toint(self.get_bytes(1))
                minuto = self.toint(self.get_bytes(1))
                segundo = self.toint(self.get_bytes(1))
                dia = self.toint(self.get_bytes(1)) + 1
                mes = self.toint(self.get_bytes(1))
                ano = self.toint(self.get_bytes(1)) + 2000
                hora = datetime(ano, mes, dia, hora, minuto, segundo)
                # tiempo
                ts = [None] * 5
                for i in range(5):
                    us = self.get_bytes(4)
                    ts[i] = us
                muestras = self.get_bytes(1)
                metodo = self.metodos[self.get_bytes(1)]
                #velocidad = self.get_bytes(2)
                fila = [lote, hora] + ts + [muestras, metodo, velocidad]
                csv.append(fila)
            except:
                self.label_update('Registro con errores')
                time.sleep(0.5)
                print ano, mes, dia, hora, minuto, segundo
                print lote, hora,minuto, segundo, dia, mes, ano
                raise
            print 'fila', n , csv
            n += 1
        return csv

    def escaneando(self, num_ports=20):
        #-- Lista de los dispositivos serie. Inicialmente vacia
        dispositivos_serie = []
        self.label_update('Buscando puertos')
        print os.name
        if os.name == 'nt':
            for i in range(num_ports):
                try:
                    s = Serial(port=i, baudrate=9600, timeout=1)
                    dispositivos_serie.append(s.portstr)
                    s.close()
                except:
                    pass
        else:
            for i in range(num_ports):
                try:
                    s = Serial(port='/dev/ttyUSB%d' % i,
                        baudrate=9600, timeout=1)
                    dispositivos_serie.append('/dev/ttyUSB%d' % i)
                    s.close()
                except:
                    pass
        if len(dispositivos_serie) == 0:
            self.label_update('No hay puertos')
            print 'No hay puertos'
        else:
            for puerto in dispositivos_serie:
                s = Serial(port=puerto, baudrate=9600, timeout=2)
                com = 'AT+MD?' + chr(0x0D) + chr(0x0A)
                if os.name == 'nt':
                    self.label_update('Buscando en puerto %s' % puerto)
                else:
                    self.label_update('Buscando en puerto /dev/ttyUSB%s' % puerto)
                try:
                    s.write(com)
                    respuesta = s.read(1101)
                    if 'MIT03' in respuesta:
                        self.label_update('OK - Conectado a ' + puerto)
                        self.escanear = False
                        self.puerto_nombre = puerto
                        self.button.update('download.png', 'Descargar')
                        self.hora.set_sensitive(True)
                        self.borrar_but.set_sensitive(True)
                        s.close()
                        return
                except:
                    s.close()
        self.label_update('MIT no encontrado.')

    def quit(self, *args):
        com = 'AT+CLOSE' + chr(0x0D) + chr(0x0A)
        try:
            self.puerto.write(com)
        except:
            self.label_update('No se pudo terminar la conexión')
        self.hilogtk.quit = True
        gtk.main_quit()


class Button(gtk.Button):

    def __init__(self, path, string=None):
        super(Button, self).__init__()
        self.hbox = gtk.HBox(False, 0)
        if not path is None:
            self.imagen = gtk.Image()
            path = os.path.join('images', path)
            self.imagen.set_from_file(path)
            self.hbox.pack_start(self.imagen, True, True, 0)
        if not string is None:
            self.label = gtk.Label(string)
            self.label.set_use_underline(True)
            self.hbox.pack_start(self.label, True, True, 0)
        self.add(self.hbox)

    def update(self, path, string=None):
        self.hbox = gtk.HBox(False, 0)
        if not path is None:
            path = os.path.join('images', path)
            self.imagen.set_from_file(path)
        if not string is None:
            self.label.set_label(string)

    def vertical(self):
        self.hbox.set_orientation(gtk.ORIENTATION_VERTICAL)


class Dialogo(gtk.Dialog):

    def __init__(self, parent, titulo, imagen, mensaje, default=True):
        super(Dialogo, self).__init__(
            flags=gtk.DIALOG_MODAL | gtk.DIALOG_DESTROY_WITH_PARENT)
        self.set_modal(True)
        self.set_transient_for(parent)
        self.set_default_size(200, 50)
        self.set_title(titulo)
        self.set_position(gtk.WIN_POS_CENTER)
        self.connect("delete_event", self.cerrar)
        hbox = gtk.HBox(False, 0)
        self.vbox.pack_start(hbox, False, False, 10)
        image = gtk.Image()
        path = os.path.join('images', imagen)
        image.set_from_file(path)
        hbox.pack_start(image, False, False, 5)
        label = gtk.Label()
        label.set_markup(mensaje)
        hbox.pack_start(label, False, False, 5)
        self.but_ok = Button('ok.png', "_OK")
        but_salir = Button('no.png', "_No")
        self.add_action_widget(self.but_ok, gtk.RESPONSE_OK)
        self.add_action_widget(but_salir, gtk.RESPONSE_CANCEL)
        if default:
            self.set_focus(self.but_ok)
        else:
            self.set_focus(but_salir)

    def iniciar(self):
        self.show_all()
        if self.run() == gtk.RESPONSE_OK:
            return True
        else:
            return False

    def cerrar(self, *args):
        self.destroy()


class Alerta(gtk.Dialog):

    def __init__(self, parent, titulo, imagen, mensaje):
        super(Alerta, self).__init__(
            flags=gtk.DIALOG_MODAL | gtk.DIALOG_DESTROY_WITH_PARENT)
        self.set_modal(True)
        self.set_transient_for(parent)
        self.set_default_size(200, 50)
        self.set_title(titulo)
        self.set_position(gtk.WIN_POS_CENTER)
        self.connect("delete_event", self.cerrar)
        hbox = gtk.HBox(False, 0)
        self.vbox.pack_start(hbox, False, False, 10)
        image = gtk.Image()
        path = os.path.join('images', imagen)
        image.set_from_file(path)
        hbox.pack_start(image, False, False, 5)
        label = gtk.Label()
        label.set_markup(mensaje)
        hbox.pack_start(label, False, False, 5)
        self.but_salir = Button('ok.png', "_Ok")
        self.add_action_widget(self.but_salir, gtk.RESPONSE_CANCEL)
        self.set_focus(self.but_salir)
        self.iniciar()

    def iniciar(self):
        self.show_all()
        self.run()
        self.cerrar()

    def cerrar(self, *args):
        self.destroy()

gtk.threads_init()
gtk.threads_enter()
v = Ventana()

#respuesta = '08C2016613055903041301010101010101020000000001010104010101050502'.decode('hex')
#v.prueba_xls(respuesta)

gtk.main()
v.hilogtk.quit = True
v.hilogtk.join()
gtk.threads_leave()
if os.name == 'nt':
    os.system('taskkill /im MIT04.exe /f')
