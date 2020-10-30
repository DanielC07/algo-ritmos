import time
import win32com.client
import datetime
import random
import json
import os

qinfo=win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
computer_name = os.getenv('COMPUTERNAME')
qinfo.FormatName="direct=os:"+computer_name+"\\PRIVATE$\\COLA_CAJ" 
sac=win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
sac.FormatName="direct=os:"+computer_name+"\\PRIVATE$\\COLA_SAC"

cont_caj = 0
cont_sac = 0

lista_caj = []
lista_sac = []

def enviar(x):
    queue=qinfo.Open(2,0)
    msg=win32com.client.Dispatch("MSMQ.MSMQMessage")
    msg.Label="TicketNo."
    msg.Body = x
    msg.Send(queue)
    queue.Close()

def enviar2(x):
    queue=sac.Open(2,0)
    msg=win32com.client.Dispatch("MSMQ.MSMQMessage")
    msg.Label="TicketNo."
    msg.Body = x
    msg.Send(queue)
    queue.Close()

def relleno(w):
    r = ""
    if len(w)<4:
        f = 4 - len(w)
        for x in range(f):
            r += "0"
        r += str(w)
        return r
    else: 
        return w

def generarTicket(x,fecha):
    if x == "1":
        a = "SAC-"
    elif x == "2":
        a = "CAJ-"

    numero = str(random.randint(1,1000))
    numero = relleno(numero)
    fecha2 = format(fecha.year) + format(fecha.month) + format(fecha.day)
    ticket = a  + fecha2 + numero
    
    return ticket

def recibe():
    queue=qinfo.Open(1,0)
    msg=queue.Receive()
    queue.Close()

def recibe2():
    queue=sac.Open(1,0)
    msg=queue.Receive()
    queue.Close()

class Base:
    def _init_(self):
        self.ticket = ""
        self.agente = ""
        self.date = ""
        self.tipooperacion = None       

while True:
    os.system('cls')
    print("""=======================
= ALACOLAPP GENERATOR =
=======================
1. Generar Ticket
2. Control de cola
3. Llamada de cliente
4. Agente de operacion
5. Salir
    """)
    try:
        op = int(input("Ingresa una opcion disponible: "))
        if op == 1:
            while True:
                os.system('cls')
                print("Bienvenido!! Que deseas Hacer? \n \n1.GENERAR NUEVO TICKET \n2.REGRESAR")
                opc = input()
                if opc.strip() == "1":
                    ticket = Base()
                    while True:
                        os.system('cls')
                        print(" \nEl Ticket que usted desea generar es para: \n1. SAC \n2. CAJ")
                        x = input()
                        if x.strip() == "1" or x.strip() == "2":
                            break
                        else:
                            print(" \nLa opcion Ingresada No existe, Por Favor Presiones Enter e Intentelo Nuevamente")
                            input()
                    fec = datetime.datetime.now()
                    fecha = fec.strftime('%Y/%m/%d %H:%M:%S')
                    ticket.ticket = generarTicket(x,fec)
                    ticket.date = str(fecha)
                    ticket.tipooperacion = ticket.ticket[0:3]
                    ticket.agente = None 
                    print(" \nFelicitaciones! Su Ticket a sido generado con exito! : " + ticket.ticket)
                    ticketjson = json.dumps(ticket.__dict__)
                    input()
                    while True:
                        os.system('cls')
                        print(" \nDesea Encolar el Ticket Generado? \n1. SI \n2. NO")
                        e = input()
                        if e.strip() == "1" or e.strip() == "2":
                            if e == "1":
                                print(" \nEncolando Ticket...")
                                time.sleep(2)
                                print(" \nEl Ticket fue encolado de manera exitosa! ")
                                input()
                                print(" \nRegresando al Menú Principal...")
                                if ticket.ticket[0:3] == "SAC":
                                    lista_sac.append(ticket.ticket)
                                    enviar2(ticketjson)
                                    cont_sac += 1
                                else:
                                    lista_caj.append(ticket.ticket)
                                    enviar(ticketjson)
                                    cont_caj += 1
                                time.sleep(2)
                            elif e == "2":
                                print(" \nRegresando al Menú Principal...")
                                time.sleep(2)
                            else:
                                print(" \nLa opcion Ingresada No existe, Por Favor Presiones Enter e Intentelo Nuevamente")
                                input()
                            break
                        else:
                            print(" \nLa opcion Ingresada No existe, Por Favor Presiones Enter e Intentelo Nuevamente")
                            input()
                    

                elif opc.strip() == "2":
                    print(" \nGracias por Utilizar Mi Programa! Vuelva Pronto!")
                    input()
                    break
                else:
                    print(" \nLa opcion Ingresada No existe, Por Favor Presiones Enter e Intentelo Nuevamente")
                    input()
        elif op == 3:
            if cont_caj == 0 and cont_sac == 0:
                print("No hay clientes en espera\n")
                input()
            else:
                if(cont_caj > 0):
                    print(f"{lista_caj[0]} Por favor pasar a la ventanilla 1")
                if cont_sac > 0:
                    print(f"{lista_sac[0]} Por favor pasar a la ventanilla 2")
                input()


        elif op == 4:
                while True:
                    print("""=======================
= AGENTE DE OPERACION =
=======================
1. Atender a Cliente
2. Regresar
    """)
                    op = int(input("Ingresa una opcion disponible: "))
                    if (op == 1):
                        if cont_caj == 0 and cont_sac == 0:
                            print("No hay clientes en espera\n")
                        else:
                            if(cont_caj > 0):
                                lista_caj.pop(0)
                                print("Ventanilla: 1")
                                print("Tipo operacion: CAJ")
                                recibe()
                                cont_caj -= 1
                            else:
                                print("Ventanilla 1 está Vacia, No hay clientes en la cola CAJ")
                            if cont_sac > 0:
                                print("Ventanilla: 2")
                                print("Tipo operacion: SAC")

                                lista_sac.pop(0)
                                recibe2()
                                cont_sac -= 1
                            else:
                                print("Ventanilla 2 está Vacia, No hay clientes en la cola SAC")
                    elif (op == 2):
                        break
                    else:
                       print(" \nLa opcion Ingresada No existe, Por Favor Presiones Enter e Intentelo Nuevamente")
        elif op == 5:
            break
        else:
            print("Opcion no disponible\n")
    except ValueError:
        print("No puede ingresar caracteres en esta opción\n")