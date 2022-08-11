from cProfile import label
from tkinter import *
from turtle import width
from typing import final
from PIL import Image, ImageTk
import tkinter.messagebox as tkmsg
import mysql.connector as mysql
from tkinter import ttk
import smtplib
from email.message import EmailMessage
import random
import pandas as pd
from win32com import client

root = Tk()
root.geometry("1540x800+0+0")
root.title("One stop bike care")


def inext():
    new_window = Toplevel(root)
    new_window.title("User Info")
    new_window.geometry("1540x800+0+0")
    new_window.config(background="powder blue")

    Mainframe = Frame(new_window)
    Mainframe.grid()

    top = Frame(Mainframe, bd = 14, width=1540, height=600, padx=20, relief=RIDGE, bg = "cadet blue")
    top.pack(side=TOP)

    left = Frame(top, bd = 10, width=580, height=600, padx=2, relief=RIDGE, bg = "powder blue")
    left.pack(side=LEFT)

    right = Frame(top, bd = 10, width=900, height=600, padx=2, relief=RIDGE, bg = "cadet blue")
    right.pack(side=RIGHT)

    bottom = Frame(Mainframe, bd = 10, width=1500, height=80, padx=20, relief=RIDGE, bg = "powder blue")
    bottom.pack(side=BOTTOM)

    bottom1 = Frame(bottom, bd = 10, width=1500, height=80, padx=20, relief=RIDGE, bg = "powder blue")
    bottom1.pack(side=TOP)

    bottom2 = Frame(bottom, bd = 10, width=1500, height=80, padx=20, relief=RIDGE, bg = "powder blue")
    bottom2.pack(side=BOTTOM)

    def iExit():
        iExit = tkmsg.askyesno("Confirmation", "Are you sure you want to exit ?")
        if iExit > 0:
            new_window.destroy()
            return

    def display():
        # self.txtreceipt.insert(END, n.get()+"     \t\t"+p.get()+"     \t\t"+a.get()+"         \t\t"+
        # v.get()+"   \t\t\t"+pm.get()+"     \t\t\t"+DOB.get()+"\n")
        con = mysql.connect(host="localhost", user="Nishad", passwd="renukama", database="oop")
        cursor = con.cursor()
        cursor.execute("select * from usersa")
        result = cursor.fetchall()
        if len(result) != 0:
            records.delete(*records.get_children())
            for row in result:
                records.insert('', END, values=row)
            con.commit()
        con.close()
        # tkmsg.showinfo("Entry", "Record entered successfully")

    def user_info(ev):
        info = records.focus()
        data = records.item(info)
        row = data['values']
        n.set(row[0])
        p.set(row[1])
        a.set(row[2])
        v.set(row[3])
        pm.set(row[4])
        DOB.set(row[5])

    def update():
        Name = n.get()
        Phone = p.get()
        Address = a.get()
        V_number = v.get()
        pay_mode = pm.get()
        dob = DOB.get()
        
        con = mysql.connect(host="localhost", user="Nishad", passwd="renukama", database="oop")
        cursor = con.cursor()
        cursor.execute("update usersa set Phone=%s, Address=%s, V_number=%s, pay_mode=%s, dob=%s where Name=%s", (
        Phone, Address, V_number, pay_mode, dob, Name
        ))
        con.commit()
        display()
        con.close()
        
        tkmsg.showinfo("Status", "Record updated successfully")

    def delete():
        Name = n.get()
        Phone = p.get()
        Address = a.get()
        V_number = v.get()
        pay_mode = pm.get()
        dob = DOB.get()
        SR = "1"
        
        
        con = mysql.connect(host="localhost", user="Nishad", passwd="renukama", database="oop")
        cursor = con.cursor()
        sql = "delete from usersa where Name=%s"
        adr = (Name, )
        cursor.execute(sql, adr)
        con.commit()
        display()
        con.commit()
        con.close()
        tkmsg.showinfo("Status", "Record deleted successfully")

        reset()

    def search():
        Name = n.get()
        Phone = p.get()
        Address = a.get()
        V_number = v.get()
        pay_mode = pm.get()
        dob = DOB.get()


        try:
            con = mysql.connect(host="localhost", user="Nishad", passwd="renukama", database="oop")
            cursor = con.cursor()
            sql = "delete from usersa where Name=%s"
            adr = (Name, )
            # cursor.execute("select * from usersa where SR=%s", SR)
            cursor.execute(sql, adr)
            row = cursor.fetchall()
            con.commit()
            n.set(row[0])
            p.set(row[1])
            a.set(row[2])
            v.set(row[3])
            pm.set(row[4])
            DOB.set(row[5])
            con.commit()

        except:
            tkmsg.showinfo("Status", "Not found")
            reset()
        con.close()

    def reset():
        n.set("")
        p.set("")
        a.set("")
        v.set("")
        pm.set("")
        DOB.set("")
        # self.txtreceipt.delete("1.0", END)

    def submit():
        Name = n.get()
        Phone = p.get()
        Address = a.get()
        V_number = v.get()
        pay_mode = pm.get()
        dob = DOB.get()

        con = mysql.connect(host="localhost", user="Nishad", passwd="renukama", database="oop")
        cursor = con.cursor()
        # cursor.execute("insert into customer values ('"+ Name +"', '"+ Phone +"', '"+ Address +"', '"+ V_number +"', '"+ pay_mode +"', '"+ dob +"')")
        cursor.execute("insert into usersa values (%s, %s, %s, %s, %s, %s)", (
        Name, Phone, Address, V_number, pay_mode, dob
        ))
        # cursor.execute("commit")
        con.commit()
        con.close()
        
        tkmsg.showinfo("Status", "Record saved successfully")



    def donext():
        new_window3 = Toplevel(new_window)
        new_window3.geometry("1540x800+0+0")
        new_window3.title("Servive form")
        new_window3.configure(background='powder blue')
        Tops = Frame(new_window3, bg='light blue',bd=20,pady=5,relief=RIDGE)
        Tops.pack(side=TOP)

        lblTitle=Label(Tops,font=('arial',40,'bold'),text='Service form',bd=10,bg='black',
                        fg='cornsilk',justify=CENTER)
        lblTitle.grid(row=0)


        ReceiptCal_F = Frame(new_window3,bg='light blue',bd=10,relief=RIDGE)
        ReceiptCal_F.pack(side=RIGHT)

        Buttons_F=Frame(ReceiptCal_F,bg='light blue',bd=3,relief=RIDGE)
        Buttons_F.pack(side=BOTTOM)

        # Cal_F=Frame(ReceiptCal_F,bg='light blue',bd=6,relief=RIDGE)
        # Cal_F.pack(side=TOP)

        # Receipt_F=Frame(ReceiptCal_F,bg='light blue',bd=4,relief=RIDGE)
        # Receipt_F.pack(side=BOTTOM)

        # scroolbar = Scrollbar(root, orient=VERTICAL)
        # scroolbar.pack(side=RIGHT, fill=Y)
        Receipt_F=Frame(ReceiptCal_F,bg='light blue',bd=7,relief=RIDGE)
        Receipt_F.pack(side=TOP)

        MenuFrame = Frame(new_window3,bg='light blue',bd=10,relief=RIDGE)
        MenuFrame.pack(side=LEFT)
        Cost_F=Frame(MenuFrame,bg='light blue',bd=4)
        Cost_F.pack(side=BOTTOM)

        separate_F=Frame(MenuFrame,bg='light blue',bd=4)
        separate_F.pack(side=TOP)

        mail = StringVar()
        lblname = Label(separate_F, font=('arial', 12, 'bold'), text="Mail", bg="powder blue")
        lblname.grid(row=0, column=0, sticky=W)
        txtname = Entry(separate_F, font=('arial', 12, 'bold'), textvariable=mail, width=20)
        txtname.grid(row=0, column=1, padx=3, pady=20)

        cn = StringVar()
        lblname = Label(separate_F, font=('arial', 12, 'bold'), text="Contact No", bg="powder blue")
        lblname.grid(row=0, column=2, sticky=W)
        txtname = Entry(separate_F, font=('arial', 12, 'bold'), textvariable=cn, width=20)
        txtname.grid(row=0, column=3, padx=3, pady=20)


        inspect_F=Frame(MenuFrame,bg='light blue',bd=4,relief=RIDGE)
        inspect_F.pack(side=LEFT)
        replace_F=Frame(MenuFrame,bg='light blue',bd=4,relief=RIDGE)
        replace_F.pack(side=RIGHT)


        var1=IntVar()
        var2=IntVar()
        var3=IntVar()
        var4=IntVar()
        var5=IntVar()
        var6=IntVar()
        var7=IntVar()
        var8=IntVar()
        var9=IntVar()
        var10=IntVar()
        var11=IntVar()
        var12=IntVar()
        var13=IntVar()
        var14=IntVar()
        var15=IntVar()
        var16=IntVar()

        DateofOrder = StringVar()
        Receipt_Ref = StringVar()
        PaidTax = StringVar()
        SubTotal = StringVar()
        TotalCost = StringVar()
        Costofreplacement = StringVar()
        Costofinspection = StringVar()
        ServiceCharge = StringVar()

        text_Input = StringVar()
        operator = ""

        E_Engine_oil = StringVar(new_window3, value='200')
        E_oil_filter = StringVar()
        E_spark_plug = StringVar()
        E_air_filter = StringVar()
        E_cvt_filter = StringVar()
        E_drive_belt = StringVar()
        E_cvt_rollers = StringVar()
        E_hose_fuel = StringVar()

        E_clutch_shoes = StringVar()
        E_front_suspension = StringVar()
        E_control_cables = StringVar()
        E_brake_fluid = StringVar()
        E_brake_hose = StringVar()
        E_engine_decarb = StringVar()
        E_front_wheel_bearing = StringVar()
        E_rear_wheel_bearing = StringVar()

        E_Engine_oil.set("0")
        E_oil_filter.set("0")
        E_spark_plug.set("0")
        E_air_filter.set("0")
        E_cvt_filter.set("0")
        E_drive_belt.set("0")
        E_cvt_rollers.set("0")
        E_hose_fuel.set("0")

        E_clutch_shoes.set("0")
        E_front_suspension.set("0")
        E_control_cables.set("0")
        E_brake_fluid.set("0")
        E_brake_hose.set("0")
        E_engine_decarb.set("0")
        E_front_wheel_bearing.set("0")
        E_rear_wheel_bearing.set("0")

        # DateofOrder.set(time.strftime("%d/%m/%y"))

        ##########################################Function Declaration####################################################

        def iExit():
            iExit=tkmsg.askyesno("Exit System","Confirm if you want to exit")
            if iExit > 0:
                new_window3.destroy()
                return

        def isendmail():
            Mail = mail.get()
            Engine_oil = str(var1.get())
            oil_filter = str(var2.get()) 
            spark_plug = str(var3.get())
            air_filter = str(var4.get())
            cvt_filter = str(var5.get())
            drive_belt = str(var6.get())
            cvt_rollers = str(var7.get())
            hose_fuel = str(var8.get())
            clutch_shoes = str(var9.get())
            front_suspension = str(var10.get())
            control_cables = str(var11.get())
            brake_fluid = str(var12.get())
            brake_hose = str(var13.get())
            engine_decarb = str(var14.get())
            front_wheel_bearing = str(var15.get())
            # rear_wheel_bearing = str(var16.get())

            con = mysql.connect(host="localhost", user="Nishad", passwd="renukama", database="oop")
            cursor = con.cursor()
            cursor.execute("insert into service values (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)", (
            Mail, Engine_oil, oil_filter, spark_plug, air_filter, cvt_filter, drive_belt, cvt_rollers, 
            hose_fuel, clutch_shoes, front_suspension, control_cables, brake_fluid, brake_hose, engine_decarb, 
            front_wheel_bearing
            ))
            con.commit()
            con.close()

            sender_mail = "nispk1506@gmail.com"
            receiver_mail = mail.get()
            password = "renukama"

            msg = EmailMessage()
            msg['Subject'] = 'Service info'
            msg['From'] = sender_mail
            msg['To'] = receiver_mail
            msg.set_content('Hello !\n The attached file contains details of your selected services\n\n')

            files = ['info.txt']

            for file in files:
                with open(file, 'rb') as f:
                    file_data = f.read()
                    file_name = f"{mail.get()}.txt"

            msg.add_attachment(file_data, maintype = 'application', subtype = 'octet-stream', filename = file_name)

            with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                smtp.ehlo()
                smtp.starttls()
                smtp.ehlo()
                smtp.login(sender_mail, password)

                smtp.send_message(msg)
            tkmsg.showinfo("Result", "Mail Sent")


        def iprint():
            f = open("info.txt", "w")
            f.write("Hii\n")
            f.write(mail.get() + '\n\n')
            f.write('Engine oil:' + 'Rs ' + str(int(E_Engine_oil.get())) +'\n')
            f.write('Oil filter:'+ 'Rs ' + str(int(E_oil_filter.get())) +'\n')
            f.write('Spark plug:'+ 'Rs ' + str(int(E_spark_plug.get())) +'\n')
            f.write('Air filter:'+ 'Rs ' + str(int(E_air_filter.get())) +'\n')
            f.write('CVT filter:'+ 'Rs ' + str(int(E_cvt_filter.get())) +'\n')
            f.write('Drive belt:'+ 'Rs ' + str(int(E_drive_belt.get())) +'\n')
            f.write('CVT rollers:'+ 'Rs ' + str(int(E_cvt_rollers.get())) +'\n')
            f.write('Hose fuel:'+ 'Rs ' + str(int(E_hose_fuel.get()))+'\n')
            f.write('Clutch shoes:'+ 'Rs ' + str(int(E_clutch_shoes.get()))+'\n')
            f.write('Front suspension:'+ 'Rs ' + str(int(E_front_suspension.get()))+'\n')
            f.write('Control cables:'+ 'Rs ' + str(int(E_control_cables.get()))+'\n')
            f.write('Brake fluid:'+ 'Rs ' + str(int(E_brake_fluid.get()))+'\n')
            f.write('Brake hose:'+ 'Rs ' + str(int(E_brake_hose.get()))+'\n')
            f.write('Engine decarb:'+ 'Rs ' + str(int(E_engine_decarb.get()))+'\n')
            f.write('Front wheel bearing:'+ 'Rs ' + str(int(E_front_wheel_bearing.get()))+'\n')
            f.write('Rear wheel bearing:'+ 'Rs ' + str(int(E_rear_wheel_bearing.get()))+'\n\n')
            x = int(E_Engine_oil.get()) + int(E_oil_filter.get()) + int(E_spark_plug.get()) + int(E_air_filter.get()) + int(E_cvt_filter.get()) + int(E_drive_belt.get())+ int(E_cvt_rollers.get()) + int(E_hose_fuel.get()) + int(E_clutch_shoes.get())+ int(E_front_suspension.get()) + int(E_control_cables.get()) + int(E_brake_fluid.get())+ int(E_brake_hose.get()) + int(E_engine_decarb.get()) + int(E_front_wheel_bearing.get())+ int(E_rear_wheel_bearing.get())
            ser_chrg = x + 16
            tax = (x + 16)*0.15
            final = str((ser_chrg + tax)) 
            f.write('Final cost:'+'Rs '+ final + '\n')
            f.close()

            # Engine_oil = var1.get()
            # oil_filter = var2.get()
            # con = mysql.connect(host="localhost", user="Nishad", passwd="renukama", database="oop")
            # cursor = con.cursor()
            # cursor.execute("update usersa set oil_filter = %d where Engine_oil = %d", (
            # oil_filter, Engine_oil
            # ))
            # con.commit()
            # con.close()
            # tkmsg.showinfo("Status", "Information printed !")
            
        def Reset():

            PaidTax.set("")
            SubTotal.set("")
            TotalCost.set("")
            Costofreplacement.set("")
            Costofinspection.set("")
            ServiceCharge.set("")
            txtReceipt.delete("1.0",END)


            E_Engine_oil.set("0")
            E_oil_filter.set("0")
            E_spark_plug.set("0")
            E_air_filter.set("0")
            E_cvt_filter.set("0")
            E_drive_belt.set("0")
            E_cvt_rollers.set("0")
            E_hose_fuel.set("0")

            E_clutch_shoes.set("0")
            E_front_suspension.set("0")
            E_control_cables.set("0")
            E_brake_fluid.set("0")
            E_brake_hose.set("0")
            E_engine_decarb.set("0")
            E_front_wheel_bearing.set("0")
            E_rear_wheel_bearing.set("0")

            mail.set("")

            var1.set(0)
            var2.set(0)
            var3.set(0)
            var4.set(0)
            var5.set(0)
            var6.set(0)
            var7.set(0)
            var8.set(0)
            var9.set(0)
            var10.set(0)
            var11.set(0)
            var12.set(0)
            var13.set(0)
            var14.set(0)
            var15.set(0)
            var16.set(0)

            txtEngine_oil.configure(state=DISABLED)
            txtoil_filter.configure(state=DISABLED)
            txtspark_plug.configure(state=DISABLED)
            txtair_filter.configure(state=DISABLED)
            txtcvt_filter.configure(state=DISABLED)
            txtdrive_belt.configure(state=DISABLED)
            txtcvt_rollers.configure(state=DISABLED)
            txthose_fuel.configure(state=DISABLED)
            txtclutch_shoes.configure(state=DISABLED)
            txtfront_suspension.configure(state=DISABLED)
            txtcontrol_cables.configure(state=DISABLED)
            txtbrake_fluid.configure(state=DISABLED)
            txtbrake_hose.configure(state=DISABLED)
            txtengine_decarb.configure(state=DISABLED)
            txtfront_wheel_bearing.configure(state=DISABLED)
            txtrear_wheel_bearing.configure(state=DISABLED)

        def CostofItem():
            Item1=float(E_Engine_oil.get())
            Item2=float(E_oil_filter.get())
            Item3=float(E_spark_plug.get())
            Item4=float(E_air_filter.get())
            Item5=float(E_cvt_filter.get())
            Item6=float(E_drive_belt.get())
            Item7=float(E_cvt_rollers.get())
            Item8=float(E_hose_fuel.get())

            Item9=float(E_clutch_shoes.get())
            Item10=float(E_front_suspension.get())
            Item11=float(E_control_cables.get())
            Item12=float(E_brake_fluid.get())
            Item13=float(E_brake_hose.get())
            Item14=float(E_engine_decarb.get())
            Item15=float(E_front_wheel_bearing.get())
            Item16=float(E_rear_wheel_bearing.get())

            Priceofinspection =(Item1) + (Item2) + (Item3) + (Item4) + (Item5) + (Item6) + (Item7) + (Item8)

            Priceofreplacement =(Item9) + (Item10) + (Item11) + (Item12) + (Item13) + (Item14) + (Item15) + (Item16)



            inspectPrice = "₹",str('%.2f'%(Priceofinspection))
            replacePrice =  "₹",str('%.2f'%(Priceofreplacement))
            Costofreplacement.set(replacePrice)
            Costofinspection.set(inspectPrice)
            SC = "₹",str('%.2f'%(16))
            ServiceCharge.set(SC)

            SubTotalofITEMS = "₹",str('%.2f'%(Priceofinspection + Priceofreplacement + 16))
            SubTotal.set(SubTotalofITEMS)

            Tax = "₹",str('%.2f'%((Priceofinspection + Priceofreplacement + 16) * 0.15))
            PaidTax.set(Tax)

            TT=((Priceofinspection + Priceofreplacement + 16) * 0.15)
            TC="₹",str('%.2f'%(Priceofinspection + Priceofreplacement + 16 + TT))
            TotalCost.set(TC)


        def chkEngine_oil():
            if(var1.get() == 1):
                txtEngine_oil.configure(state = NORMAL)
                txtEngine_oil.focus()
                txtEngine_oil.delete('0',END)
                E_Engine_oil.set("200")
            elif(var1.get() == 0):
                txtEngine_oil.configure(state = DISABLED)
                E_Engine_oil.set("")

        def chkoil_filter():
            if(var2.get() == 1):
                txtoil_filter.configure(state = NORMAL)
                txtoil_filter.focus()
                txtoil_filter.delete('0',END)
                E_oil_filter.set("150")
            elif(var2.get() == 0):
                txtoil_filter.configure(state = DISABLED)
                E_oil_filter.set("0")

        def chk_spark_plug():
            if(var3.get() == 1):
                txtspark_plug.configure(state = NORMAL)
                txtspark_plug.focus()
                txtspark_plug.delete('0',END)
                E_spark_plug.set("130")
            elif(var3.get() == 0):
                txtspark_plug.configure(state = DISABLED)
                E_spark_plug.set("0")

        def chk_air_filter():
            if(var4.get() == 1):
                txtair_filter.configure(state = NORMAL)
                txtair_filter.focus()
                txtair_filter.delete('0',END)
                E_air_filter.set("170")
            elif(var4.get() == 0):
                txtair_filter.configure(state = DISABLED)
                E_air_filter.set("0")

        def chk_cvt_filter():
            if(var5.get() == 1):
                txtcvt_filter.configure(state = NORMAL)
                txtcvt_filter.focus()
                txtcvt_filter.delete('0',END)
                E_cvt_filter.set("170")
            elif(var5.get() == 0):
                txtcvt_filter.configure(state = DISABLED)
                E_cvt_filter.set("0")

        def chk_drive_belt():
            if(var6.get() == 1):
                txtdrive_belt.configure(state = NORMAL)
                txtdrive_belt.focus()
                txtdrive_belt.delete('0',END)
                E_drive_belt.set("75")
            elif(var6.get() == 0):
                txtdrive_belt.configure(state = DISABLED)
                E_drive_belt.set("0")

        def chk_cvt_rollers():
            if(var7.get() == 1):
                txtcvt_rollers.configure(state = NORMAL)
                txtcvt_rollers.focus()
                txtcvt_rollers.delete('0',END)
                E_cvt_rollers.set("90")
            elif(var7.get() == 0):
                txtcvt_rollers.configure(state = DISABLED)
                E_cvt_rollers.set("0")

        def chk_hose_fuel():
            if(var8.get() == 1):
                txthose_fuel.configure(state = NORMAL)
                txthose_fuel.focus()
                txthose_fuel.delete('0',END)
                E_hose_fuel.set("182")
            elif(var8.get() == 0):
                txthose_fuel.configure(state = DISABLED)
                E_hose_fuel.set("0")

        def chk_clutch_shoes():
            if(var9.get() == 1):
                txtclutch_shoes.configure(state = NORMAL)
                txtclutch_shoes.focus()
                txtclutch_shoes.delete('0',END)
                E_clutch_shoes.set("82")
            elif(var9.get() == 0):
                txtclutch_shoes.configure(state = DISABLED)
                E_clutch_shoes.set("0")

        def chk_front_suspension():
            if(var10.get() == 1):
                txtfront_suspension.configure(state = NORMAL)
                txtfront_suspension.focus()
                txtfront_suspension.delete('0',END)
                E_front_suspension.set("120")
            elif(var10.get() == 0):
                txtfront_suspension.configure(state = DISABLED)
                E_front_suspension.set("0")

        def chk_control_cables():
            if(var11.get() == 1):
                txtcontrol_cables.configure(state = NORMAL)
                txtcontrol_cables.focus()
                txtcontrol_cables.delete('0',END)
                E_control_cables.set("165")
            elif(var11.get() == 0):
                txtcontrol_cables.configure(state = DISABLED)
                E_control_cables.set("0")

        def chk_brake_fluid():
            if(var12.get() == 1):
                txtbrake_fluid.configure(state = NORMAL)
                txtbrake_fluid.focus()
                txtbrake_fluid.delete('0',END)
                E_brake_fluid.set("85")
            elif(var12.get() == 0):
                txtbrake_fluid.configure(state = DISABLED)
                E_brake_fluid.set("0")

        def chk_brake_hose():
            if(var13.get() == 1):
                txtbrake_hose.configure(state = NORMAL)
                txtbrake_hose.focus()
                txtbrake_hose.delete('0',END)
                E_brake_hose.set("90")
            elif(var13.get() == 0):
                txtbrake_hose.configure(state = DISABLED)
                E_brake_hose.set("0")

        def chk_engine_decarb():
            if(var14.get() == 1):
                txtengine_decarb.configure(state = NORMAL)
                txtengine_decarb.focus()
                txtengine_decarb.delete('0',END)
                E_engine_decarb.set("220")
            elif(var14.get() == 0):
                txtengine_decarb.configure(state = DISABLED)
                E_engine_decarb.set("0")

        def chk_front_wheel_bearing():
            if(var15.get() == 1):
                txtfront_wheel_bearing.configure(state = NORMAL)
                txtfront_wheel_bearing.focus()
                txtfront_wheel_bearing.delete('0',END)
                E_front_wheel_bearing.set("350")
            elif(var15.get() == 0):
                txtfront_wheel_bearing.configure(state = DISABLED)
                E_front_wheel_bearing.set("0")

        def chk_rear_wheel_bearing():
            if(var16.get() == 1):
                txtrear_wheel_bearing.configure(state = NORMAL)
                txtrear_wheel_bearing.focus()
                txtrear_wheel_bearing.delete('0',END)
                E_rear_wheel_bearing.set("350")
            elif(var16.get() == 0):
                txtrear_wheel_bearing.configure(state = DISABLED)
                E_rear_wheel_bearing.set("0")

        def Receipt():
            txtReceipt.delete("1.0",END)
            x=random.randint(10908,500876)
            randomRef= str(x)
            Receipt_Ref.set("Bill"+ randomRef)
            # txtReceipt.insert(END,'Receipt Ref:\t\t\t'+Receipt_Ref.get() +'\t'+ DateofOrder.get() +'\n')
            txtReceipt.insert(END,'Receipt Ref:\t\t\t'+Receipt_Ref.get() +'\n')
            txtReceipt.insert(END,'Items\t\t\t\t'+"Total Cost\n")
            txtReceipt.insert(END,'Engine oil:\t\t\t\t\t' + '₹' + str(int(E_Engine_oil.get())) +'\n')
            txtReceipt.insert(END,'Oil filter:\t\t\t\t\t'+ '₹' + str(int(E_oil_filter.get())) +'\n')
            txtReceipt.insert(END,'Spark plug:\t\t\t\t\t'+ '₹' + str(int(E_spark_plug.get())) +'\n')
            txtReceipt.insert(END,'Air filter:\t\t\t\t\t'+ '₹' + str(int(E_air_filter.get())) +'\n')
            txtReceipt.insert(END,'CVT filter:\t\t\t\t\t'+ '₹' + str(int(E_cvt_filter.get())) +'\n')
            txtReceipt.insert(END,'Drive belt:\t\t\t\t\t'+ '₹' + str(int(E_drive_belt.get())) +'\n')
            txtReceipt.insert(END,'CVT rollers:\t\t\t\t\t'+ '₹' + str(int(E_cvt_rollers.get())) +'\n')
            txtReceipt.insert(END,'Hose fuel:\t\t\t\t\t'+ '₹' + str(int(E_hose_fuel.get()))+'\n')
            txtReceipt.insert(END,'Clutch shoes:\t\t\t\t\t'+ '₹' + str(int(E_clutch_shoes.get()))+'\n')
            txtReceipt.insert(END,'Front suspension:\t\t\t\t\t'+ '₹' + str(int(E_front_suspension.get()))+'\n')
            txtReceipt.insert(END,'Control cables:\t\t\t\t\t'+ '₹' + str(int(E_control_cables.get()))+'\n')
            txtReceipt.insert(END,'Brake fluid:\t\t\t\t\t'+ '₹' + str(int(E_brake_fluid.get()))+'\n')
            txtReceipt.insert(END,'Brake hose:\t\t\t\t\t'+ '₹' + str(int(E_brake_hose.get()))+'\n')
            txtReceipt.insert(END,'Engine decarb:\t\t\t\t\t'+ '₹' + str(int(E_engine_decarb.get()))+'\n')
            txtReceipt.insert(END,'Front wheel bearing:\t\t\t\t\t'+ '₹' + str(int(E_front_wheel_bearing.get()))+'\n')
            txtReceipt.insert(END,'Rear wheel bearing:\t\t\t\t\t'+ '₹' + str(int(E_rear_wheel_bearing.get()))+'\n')
            txtReceipt.insert(END,'Cost of inspection:\t\t\t\t'+ Costofinspection.get()+'\nTax Paid:\t\t\t\t'+PaidTax.get()+"\n")
            txtReceipt.insert(END,'Cost of replacement:\t\t\t\t'+ Costofreplacement.get()+'\nSubTotal:\t\t\t\t'+str(SubTotal.get())+"\n")
            txtReceipt.insert(END,'Service Charge:\t\t\t\t'+ ServiceCharge.get()+'\nTotal Cost:\t\t\t\t'+str(TotalCost.get())+"\n")

        
        def gonext():
            new_window4 = Toplevel(new_window3)
            new_window4.geometry("1540x800+0+0")
            new_window4.title("Thank you")

            def get():
                tkmsg.showinfo("FINAL", f"{slide1.get()} star rating")

            def goexit():
                goExit=tkmsg.askyesno("Exit System","Confirm if you want to exit")
                if goExit > 0:
                    new_window4.destroy()
                    return

            # image1 = Image.open("thx.jpg")
            # photo1 = ImageTk.PhotoImage(image1)
            # labelt = Label(image=photo1)
            # labelt.pack()
            frm = Frame(new_window4, bg='light blue',bd=10,relief=RIDGE)
            frm.pack(side=BOTTOM)
            lbltxt = Label(frm, font=('arial',14,'bold'),text='Kindly Rate Us',bg='light blue',
                fg='black',justify=CENTER)
            lbltxt.grid(row=0,column=0,sticky=W)
            slide1 = Scale(new_window4, from_=1, to=5, orient=HORIZONTAL)
            slide1.pack()
            Button(new_window4, text="RATE", command=get).pack()
            Button(new_window4, text="EXIT", command=goexit).pack()

        def isubmit():
            dataframe1 = pd.read_csv("info.txt")
            # storing this dataframe in a csv file
            dataframe1.to_csv('customer.csv', index = None)
            new = pd.read_csv('customer.csv')
            exfile = pd.ExcelWriter('customer.xlsx')
            new.to_excel(exfile, index=False)
            exfile.save()
            excel = client.Dispatch("Excel.Application")
            excel.interactive = False
            excel.visible = False
            sheets = excel.Workbooks.Open('C:\\Users\\HP\\Documents\\COEP\\Python programming\\customer.xlsx')
            sheets.Activesheet.ExportAsFixedFormat(0, 'C:\\Users\\HP\\Documents\\COEP\\Python programming\\customer.xlsx')
            sheets.Close()

        Engine_oil=Checkbutton(inspect_F,text='Engine oil',variable=var1,onvalue=1,offvalue=0,font=('arial',18,'bold'),
                            bg='light blue',command=chkEngine_oil).grid(row=0,sticky=W)
        oil_filter=Checkbutton(inspect_F,text='Oil filter',variable=var2,onvalue=1,offvalue=0,font=('arial',18,'bold'),
                            bg='light blue',command=chkoil_filter).grid(row=1,sticky=W)
        spark_plug=Checkbutton(inspect_F,text='Spark plug',variable=var3,onvalue=1,offvalue=0,font=('arial',18,'bold'),
                            bg='light blue',command=chk_spark_plug).grid(row=2,sticky=W)
        air_filter=Checkbutton(inspect_F,text='Air filter',variable=var4,onvalue=1,offvalue=0,font=('arial',18,'bold'),
                            bg='light blue',command=chk_air_filter).grid(row=3,sticky=W)
        cvt_filter=Checkbutton(inspect_F,text='CVT filter',variable=var5,onvalue=1,offvalue=0,font=('arial',18,'bold'),
                            bg='light blue',command=chk_cvt_filter).grid(row=4,sticky=W)
        drive_belt=Checkbutton(inspect_F,text='Drive belt',variable=var6,onvalue=1,offvalue=0,font=('arial',18,'bold'),
                            bg='light blue',command=chk_drive_belt).grid(row=5,sticky=W)
        cvt_rollers=Checkbutton(inspect_F,text='CVT rollers',variable=var7,onvalue=1,offvalue=0,font=('arial',18,'bold'),
                            bg='light blue',command=chk_cvt_rollers).grid(row=6,sticky=W)
        hose_fuel=Checkbutton(inspect_F,text='Hose fuel',variable=var8,onvalue=1,offvalue=0,font=('arial',18,'bold'),
                            bg='light blue',command=chk_hose_fuel).grid(row=7,sticky=W)


        txtEngine_oil = Entry(inspect_F,font=('arial',16,'bold'),bd=8,width=6,justify=LEFT,state=DISABLED
                                ,textvariable=E_Engine_oil)
        txtEngine_oil.grid(row=0,column=1)

        txtoil_filter = Entry(inspect_F,font=('arial',16,'bold'),bd=8,width=6,justify=LEFT,state=DISABLED
                                ,textvariable=E_oil_filter)
        txtoil_filter.grid(row=1,column=1)

        txtspark_plug = Entry(inspect_F,font=('arial',16,'bold'),bd=8,width=6,justify=LEFT,state=DISABLED
                                ,textvariable=E_spark_plug)
        txtspark_plug.grid(row=2,column=1)

        txtair_filter= Entry(inspect_F,font=('arial',16,'bold'),bd=8,width=6,justify=LEFT,state=DISABLED
                                ,textvariable=E_air_filter)
        txtair_filter.grid(row=3,column=1)

        txtcvt_filter = Entry(inspect_F,font=('arial',16,'bold'),bd=8,width=6,justify=LEFT,state=DISABLED
                                ,textvariable=E_cvt_filter)
        txtcvt_filter.grid(row=4,column=1)

        txtdrive_belt = Entry(inspect_F,font=('arial',16,'bold'),bd=8,width=6,justify=LEFT,state=DISABLED
                                ,textvariable=E_drive_belt)
        txtdrive_belt.grid(row=5,column=1)

        txtcvt_rollers = Entry(inspect_F,font=('arial',16,'bold'),bd=8,width=6,justify=LEFT,state=DISABLED
                                ,textvariable=E_cvt_rollers)
        txtcvt_rollers.grid(row=6,column=1)

        txthose_fuel = Entry(inspect_F,font=('arial',16,'bold'),bd=8,width=6,justify=LEFT,state=DISABLED
                                ,textvariable=E_hose_fuel)
        txthose_fuel.grid(row=7,column=1)

        clutch_shoes = Checkbutton(replace_F,text="Clutch shoes",variable=var9,onvalue = 1, offvalue=0,
                                font=('arial',16,'bold'),bg='light blue',command=chk_clutch_shoes).grid(row=0,sticky=W)
        front_suspension = Checkbutton(replace_F,text="Front suspension",variable=var10,onvalue = 1, offvalue=0,
                                font=('arial',16,'bold'),bg='light blue',command=chk_front_suspension).grid(row=1,sticky=W)
        control_cables = Checkbutton(replace_F,text="Control cables",variable=var11,onvalue = 1, offvalue=0,
                                font=('arial',16,'bold'),bg='light blue',command=chk_control_cables).grid(row=2,sticky=W)
        brake_fluid = Checkbutton(replace_F,text="Brake fluid",variable=var12,onvalue = 1, offvalue=0,
                                font=('arial',16,'bold'),bg='light blue',command=chk_brake_fluid).grid(row=3,sticky=W)
        brake_hose = Checkbutton(replace_F,text="Brake hose",variable=var13,onvalue = 1, offvalue=0,
                                font=('arial',16,'bold'),bg='light blue',command=chk_brake_hose).grid(row=4,sticky=W)
        engine_decarb = Checkbutton(replace_F,text="Engine decarb",variable=var14,onvalue = 1, offvalue=0,
                                font=('arial',16,'bold'),bg='light blue',command=chk_engine_decarb).grid(row=5,sticky=W)
        front_wheel_bearing = Checkbutton(replace_F,text="Front wheel bearing",variable=var15,onvalue = 1, offvalue=0,
                                font=('arial',16,'bold'),bg='light blue',command=chk_front_wheel_bearing).grid(row=6,sticky=W)
        rear_wheel_bearing = Checkbutton(replace_F,text="Rear wheel bearing",variable=var16,onvalue = 1, offvalue=0,
                                font=('arial',16,'bold'),bg='light blue',command=chk_rear_wheel_bearing).grid(row=7,sticky=W)


        txtclutch_shoes=Entry(replace_F,font=('arial',16,'bold'),bd=8,width=6,justify=LEFT,state=DISABLED,
                                textvariable=E_clutch_shoes)
        txtclutch_shoes.grid(row=0,column=1)

        txtfront_suspension=Entry(replace_F,font=('arial',16,'bold'),bd=8,width=6,justify=LEFT,state=DISABLED,
                                textvariable=E_front_suspension)
        txtfront_suspension.grid(row=1,column=1)

        txtcontrol_cables=Entry(replace_F,font=('arial',16,'bold'),bd=8,width=6,justify=LEFT,state=DISABLED,
                                textvariable=E_control_cables)
        txtcontrol_cables.grid(row=2,column=1)

        txtbrake_fluid=Entry(replace_F,font=('arial',16,'bold'),bd=8,width=6,justify=LEFT,state=DISABLED,
                                textvariable=E_brake_fluid)
        txtbrake_fluid.grid(row=3,column=1)

        txtbrake_hose=Entry(replace_F,font=('arial',16,'bold'),bd=8,width=6,justify=LEFT,state=DISABLED,
                                textvariable=E_brake_hose)
        txtbrake_hose.grid(row=4,column=1)

        txtengine_decarb=Entry(replace_F,font=('arial',16,'bold'),bd=8,width=6,justify=LEFT,state=DISABLED,
                                textvariable=E_engine_decarb)
        txtengine_decarb.grid(row=5,column=1)

        txtfront_wheel_bearing=Entry(replace_F,font=('arial',16,'bold'),bd=8,width=6,justify=LEFT,state=DISABLED,
                                textvariable=E_front_wheel_bearing)
        txtfront_wheel_bearing.grid(row=6,column=1)

        txtrear_wheel_bearing=Entry(replace_F,font=('arial',16,'bold'),bd=8,width=6,justify=LEFT,state=DISABLED,
                                textvariable=E_rear_wheel_bearing)
        txtrear_wheel_bearing.grid(row=7,column=1)

        lblCostofinspection=Label(Cost_F,font=('arial',14,'bold'),text='Cost of Inspection\t',bg='light blue',
                        fg='black',justify=CENTER)
        lblCostofinspection.grid(row=0,column=0,sticky=W)
        txtCostofinspection=Entry(Cost_F,bg='white',bd=7,font=('arial',14,'bold'),
                                insertwidth=2,justify=RIGHT,textvariable=Costofinspection)
        txtCostofinspection.grid(row=0,column=1)

        lblCostofreplacement=Label(Cost_F,font=('arial',14,'bold'),text='Cost of Replacement',bg='light blue',
                        fg='black',justify=CENTER)
        lblCostofreplacement.grid(row=1,column=0,sticky=W)
        txtCostofreplacement=Entry(Cost_F,bg='white',bd=7,font=('arial',14,'bold'),
                                insertwidth=2,justify=RIGHT,textvariable=Costofreplacement)
        txtCostofreplacement.grid(row=1,column=1)

        lblServiceCharge=Label(Cost_F,font=('arial',14,'bold'),text='Service Charge',bg='light blue',
                        fg='black',justify=CENTER)
        lblServiceCharge.grid(row=2,column=0,sticky=W)
        txtServiceCharge=Entry(Cost_F,bg='white',bd=7,font=('arial',14,'bold'),
                                insertwidth=2,justify=RIGHT,textvariable=ServiceCharge)
        txtServiceCharge.grid(row=2,column=1)

        lblPaidTax=Label(Cost_F,font=('arial',14,'bold'),text='\tPaid Tax',bg='light blue',bd=7,
                        fg='black',justify=CENTER)
        lblPaidTax.grid(row=0,column=2,sticky=W)
        txtPaidTax=Entry(Cost_F,bg='white',bd=7,font=('arial',14,'bold'),
                                insertwidth=2,justify=RIGHT,textvariable=PaidTax)
        txtPaidTax.grid(row=0,column=3)

        lblSubTotal=Label(Cost_F,font=('arial',14,'bold'),text='\tSub Total',bg='light blue',bd=7,
                        fg='black',justify=CENTER)
        lblSubTotal.grid(row=1,column=2,sticky=W)
        txtSubTotal=Entry(Cost_F,bg='white',bd=7,font=('arial',14,'bold'),
                                insertwidth=2,justify=RIGHT,textvariable=SubTotal)
        txtSubTotal.grid(row=1,column=3)

        lblTotalCost=Label(Cost_F,font=('arial',14,'bold'),text='\tTotal',bg='light blue',bd=7,
                        fg='black',justify=CENTER)
        lblTotalCost.grid(row=2,column=2,sticky=W)
        txtTotalCost=Entry(Cost_F,bg='white',bd=7,font=('arial',14,'bold'),
                                insertwidth=2,justify=RIGHT,textvariable=TotalCost)
        txtTotalCost.grid(row=2,column=3)

        scroolbar = Scrollbar(new_window3, orient=VERTICAL)
        scroolbar.pack(side=RIGHT, fill=Y)
        txtReceipt=Text(Receipt_F,width=46,height=12,bg='white',bd=4,font=('arial',12,'bold'), yscrollcommand=scroolbar.set)
        scroolbar.config(command=txtReceipt.yview)
        txtReceipt.grid(row=0,column=0)


        btnTotal=Button(Buttons_F,padx=16,pady=1,bd=7,fg='black',font=('arial',16,'bold'),width=4,text='Total',
                                bg='light blue',command=CostofItem).grid(row=0,column=0)
        btnReceipt=Button(Buttons_F,padx=16,pady=1,bd=7,fg='black',font=('arial',16,'bold'),width=4,text='Receipt',
                                bg='light blue',command=Receipt).grid(row=0,column=1)
        btnReset=Button(Buttons_F,padx=16,pady=1,bd=7,fg='black',font=('arial',16,'bold'),width=4,text='Reset',
                                bg='light blue',command=Reset).grid(row=0,column=2)
        btnPrint=Button(Buttons_F,padx=16,pady=1,bd=7,fg='black',font=('arial',16,'bold'),width=4,text='Print',
                                bg='light blue',command=iprint).grid(row=0,column=3)
        btnExit=Button(Buttons_F,padx=16,pady=1,bd=7,fg='black',font=('arial',16,'bold'),width=4,text='Exit',
                                bg='light blue',command=iExit).grid(row=0,column=4)
        btnsendmail=Button(Buttons_F,padx=16,pady=1,bd=7,fg='black',font=('arial',16,'bold'), width=6,text='Send Mail',
        bg='light blue',command=isendmail).grid(row=1,column=1)
        btngonext=Button(Buttons_F,padx=16,pady=1,bd=7,fg='black',font=('arial',16,'bold'),width=4,text='NEXT',
                        bg='light blue', command=gonext).grid(row=1,column=3)
        btnSubmit=Button(Buttons_F,padx=16,pady=1,bd=7,fg='black',font=('arial',16,'bold'),width=4,text='Submit',
                                bg='light blue',command=isubmit).grid(row=1,column=2)


    n = StringVar()
    lblname = Label(left, font=('arial', 12, 'bold'), text="Name", bg="powder blue")
    lblname.grid(row=0, column=0, sticky=W)
    txtname = Entry(left, font=('arial', 12, 'bold'), textvariable=n, width=20)
    txtname.grid(row=0, column=1, padx=3, pady=20)

    p = StringVar()
    lblname = Label(left, font=('arial', 12, 'bold'), text="Phone Number", bg="powder blue")
    lblname.grid(row=1, column=0, sticky=W)
    txtname = Entry(left, font=('arial', 12, 'bold'), textvariable=p, width=20)
    txtname.grid(row=1, column=1, padx=3, pady=20)

    a = StringVar()
    lblname = Label(left, font=('arial', 12, 'bold'), text="Address", bg="powder blue")
    lblname.grid(row=2, column=0, sticky=W)
    txtname = Entry(left, font=('arial', 12, 'bold'), textvariable=a, width=20)
    txtname.grid(row=2, column=1, padx=3, pady=20)

    v = StringVar()
    lblname = Label(left, font=('arial', 12, 'bold'), text="Vehicle Number", bg="powder blue")
    lblname.grid(row=3, column=0, sticky=W)
    txtname = Entry(left, font=('arial', 12, 'bold'), textvariable=v, width=20)
    txtname.grid(row=3, column=1, padx=3, pady=20)

    pm = StringVar()
    lblname = Label(left, font=('arial', 12, 'bold'), text="Payment", bg="powder blue")
    lblname.grid(row=4, column=0, sticky=W)
    option = ttk.Combobox(left, textvariable=pm, state='readonly', font=('arial', 12, 'bold'),
    width=18)
    option['value'] = ('', 'Cash', 'Credit Card', 'Check', 'Online')
    option.current(0)
    option.grid(row=4, column=1, pady=3, padx=20)

    DOB = StringVar()
    lblDOB = Label(left, font=('arial', 12, 'bold'), text="Date of Birth", bg="powder blue")
    lblDOB.grid(row=5, column=0, sticky=W)
    txtDOB = Entry(left, font=('arial', 12, 'bold'), textvariable=DOB, width=20)
    txtDOB.grid(row=5, column=1, padx=3, pady=20)

    scrool_y = Scrollbar(right, orient=VERTICAL)

    records = ttk.Treeview(right, height=12, columns=("Name", "Phone", "Address", "V_no",
    "Pay", "DOB"), yscrollcommand=scrool_y.set)

    scrool_y.pack(side=RIGHT, fill=Y)

    # self.records.heading("Sr No", text="Sr No")
    records.heading("Name", text="Name")
    records.heading("Phone", text="Phone")
    records.heading("Address", text="Address")
    records.heading("V_no", text="Vehicle No")
    records.heading("Pay", text="Payment")
    records.heading("DOB", text="DOB")

    records['show'] = 'headings'

    # self.records.column("Sr No", width=40)
    records.column("Name", width=70)
    records.column("Phone", width=70)
    records.column("Address", width=100)
    records.column("V_no", width=90)
    records.column("Pay", width=70)
    records.column("DOB", width=50)

    records.pack(fill=BOTH, expand=1)
    records.bind("<ButtonRelease-1>", user_info)

    btnsubmit = Button(bottom1, padx=16, pady=1, bd=4, fg="black", font=('arial', 16, 'bold'),
    width=21, height=2, bg="powder blue", text="Submit", command=submit).grid(row=0, column=0, padx=5)

    btndisplay = Button(bottom1, padx=16, pady=1, bd=4, fg="black", font=('arial', 16, 'bold'),
    width=21, height=2, bg="powder blue", text="Display", command=display).grid(row=0, column=1, padx=5)

    btnreset = Button(bottom1, padx=16, pady=1, bd=4, fg="black", font=('arial', 16, 'bold'),
    width=21, height=2, bg="powder blue", text="Reset", command=reset).grid(row=0, column=2, padx=5)

    btnexit = Button(bottom1, padx=16, pady=1, bd=4, fg="black", font=('arial', 16, 'bold'),
    width=21, height=2, bg="powder blue", text="Exit", command=iExit).grid(row=0, column=4, padx=5)

    btnUpdate = Button(bottom2, padx=16, pady=1, bd=4, fg="black", font=('arial', 16, 'bold'),
    width=21, height=2, bg="powder blue", text="Update", command=update).grid(row=0, column=0, padx=5)

    btnDelete = Button(bottom2, padx=16, pady=1, bd=4, fg="black", font=('arial', 16, 'bold'),
    width=21, height=2, bg="powder blue", text="Delete", command=delete).grid(row=0, column=1, padx=5)

    btnSearch = Button(bottom2, padx=16, pady=1, bd=4, fg="black", font=('arial', 16, 'bold'),
    width=21, height=2, bg="powder blue", text="Search", command=search).grid(row=0, column=2, padx=5)

    label = Label(bottom2, font=('arial', 16, 'bold'), height=5, text="Please click the serive type you want on the next page", 
    bg="powder blue")
    label.grid(row=1, column=1, sticky=W)

    btnnext = Button(bottom2, padx=16, pady=1, bd=4, fg="black", font=('arial', 16, 'bold'),
    width=21, height=2, bg="powder blue", text="NEXT", command=donext).grid(row=2, column=1, padx=5)



top = Frame(root, bd = 14, width=1540, height=300, padx=20, relief=RIDGE, bg = "cadet blue")
top.pack(side=TOP)
label = Label(top, font=('arial', 65, 'bold'), text="Welcome to ONE STOP BIKE CARE", bg="powder blue")
label.grid(row=0, column=0, sticky=W)

image = Image.open("bike.jpg")
photo = ImageTk.PhotoImage(image)
label1 = Label(image=photo)
label1.pack()

bottom = Frame(root, bd = 14, width=1540, height=250, padx=20, relief=RIDGE, bg = "cadet blue")
bottom.pack(side=BOTTOM)

label = Label(bottom, font=('arial', 40, 'bold'), text="Please Enter your details on the next page", 
bg="powder blue")
label.grid(row=0, column=0, sticky=W)

btnnext = Button(bottom, padx=16, pady=1, bd=4, fg="black", font=('arial', 16, 'bold'),
        width=21, height=2, bg="powder blue", text="NEXT", command=inext).grid(row=1, column=0, padx=5)


root.mainloop()