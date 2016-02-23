#! /usr/bin/env python
#! -*- coding: utf-8 -*-

import os
from serial import Serial

num_ports = 20
#-- Lista de los dispositivos serie. Inicialmente vacia
dispositivos_serie = []
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
            s = Serial(port='/dev/ttyUSB%d' % i, baudrate=9600, timeout=1)
            dispositivos_serie.append('/dev/ttyUSB%d' % i)
            s.close()
        except:
            pass
if len(dispositivos_serie) == 0:
    print 'ERROR no hay puertos'
for puerto in dispositivos_serie:
    s = Serial(port=puerto, baudrate=9600, timeout=2)
    com = 'AT+MD?' + chr(0x0D) + chr(0x0A)
    try:
        s.write(com)
        respuesta = s.read(1101)
        if 'MIT03' in respuesta:
            com = 'AT+CLR' + chr(0x0D) + chr(0x0A)
            puerto.write(com)
            break
        else:
            s.close()
            print 'MIT-03 no encontrado'
    except:
        s.close()
