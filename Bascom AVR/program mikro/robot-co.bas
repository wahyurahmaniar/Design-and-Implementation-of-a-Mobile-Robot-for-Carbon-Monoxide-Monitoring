$regfile = "m32def.dat"                                     'mikrokontroler ATMega32
$crystal = 11059200                                         'crystal 11.0592MHz
$baud = 9600                                                'Baud Rate 9600 bps

'Deklarasi LCD
Config Lcdpin = Pin , Db4 = Portc.4 , Db5 = Portc.5 , Db6 = Portc.6 , Db7 = Portc.7 , E = Portc.3 , Rs = Portc.2
Config Lcd = 16 * 2

Deflcdchar 1 , 31 , 31 , 31 , 31 , 31 , 31 , 31 , 31

'Deklarasi ADC
Config Adc = Single , Prescaler = Auto , Reference = Avcc
Start Adc

'Variabel
Dim Yy As Word
Dim Y1 As Single , Y2 As Single , Y3 As Single
Dim Ubah_ppm As Single
Dim Ro As Single , Rs As Single
Dim Teg As Single
Dim X1 As Single , X2 As Single
Dim Ppm As Integer
Dim Ubah_teg As String * 10
Dim Datas As String * 10
Dim Pisah As Byte , Cek(3) As String * 10

Dim Batas_min As Integer , Batas_max As Integer
Dim Simpan_min As Eram Integer , Simpan_max As Eram Integer
Dim Count1 As Integer , Count2 As Integer
Dim I As Integer , J As Integer
Dim Ii As Integer , Jj As Integer
Dim A As String * 16 , B As Integer

'Konfigurasi Output
Config Portb = Output
Config Portd.5 = Output
Config Portd.6 = Output
Config Portd.7 = Output
Portb = 0
Portd.5 = 0
Portd.6 = 0
Portd.7 = 0

Ena Alias Portb.0
In1 Alias Portb.1
In2 Alias Portb.2
Enb Alias Portb.3
In3 Alias Portb.4
In4 Alias Portb.5

Hijau Alias Portd.7
Kuning Alias Portd.5
Merah Alias Portd.6

Relay1 Alias Portb.6                                        'Relay1= VH 5 Volt
Relay2 Alias Portb.7                                        'Relay2= VH 1.5 Volt

Batas_min = Simpan_min
Batas_max = Simpan_max

If Batas_min < 1 Then                                       'Kalo batas_min kosong
Batas_min = 25
Else
Batas_min = Simpan_min
End If

If Batas_max < 1 Then                                       'Kalo batas_max kosong
Batas_max = 50
Else
Batas_max = Simpan_max
End If

Cls
Cursor Off
Locate 1 , 1
Lcd "Batas min= " ; Batas_min
Locate 2 , 1
Lcd "Batas max= " ; Batas_max
Wait 1


Cls
Cursor Off
Locate 1 , 1
Lcd "Loading"
A = ""
   For B = 1 To 16
   A = A + Chr(1)
   Locate 2 , 1
   Lcd A
   Wait 1
   Next


Yy = Getadc(0)
Y1 = Yy * 5
Y2 = Y1 / 1023

'Rumus Mencari Ro atw Rs awal
'Rs= ((Vc/VRL)-1)*RL
Rs = 5 / Y2
Rs = Rs - 1
Rs = Rs * 10

'Rs/Ro= 30.34 * (1/(5^0.53))
Y3 = 5 ^ 0.53
Y3 = 1 / Y3
Y3 = 30.34 * Y3
Ro = Rs / Y3
Ppm = Int(ro)

Cls
Cursor Off
Locate 1 , 1
Lcd "Ro = " ; Ppm
Wait 1

Hijau = 1
Kuning = 0
Merah = 0

Relay1 = 1                                                  'VH=5 Volt
Relay2 = 0                                                  'VH=1.5 Volt

Count1 = 0
Count2 = 0

Ii = 1
Jj = 1

'Menampilkan pada LCD
Cls
Cursor Off
Locate 1 , 2
Lcd "Robot Pemantau"
Locate 2 , 1
Lcd "Karbon Monoksida"
Wait 1

Mulai:
Do
On Urxc Baca_serial

If Ii = 1 Then
   For I = 1 To 60                                          'Mengaktifkan VH 5 volt selama 60 detik
   Count2 = 0
   Yy = Getadc(0)                                           'ambil data ADC 0
   'Rumus untuk mengubah data ADC jadi data tegangan
   Y1 = Yy * 5
   Y2 = Y1 / 1023

   'Rs= ((Vc/VRL)-1)*RL
   Rs = 5 / Y2                                              'PPM terbaca sekarang
   Rs = Rs - 1
   Rs = Rs * 10

   Teg = Rs / Ro

   'PPM = (30.34/ tegangan) ^ (100/53)
   X1 = 30.34 / Teg
   X2 = 100 / 53
   Ubah_ppm = X1 ^ X2


   Ubah_teg = Fusing(y2 , "#.##")                           'Mengubah data tegangan menjadi string
   Ppm = Int(ubah_ppm)                                      'Mengubah data ppm jd integer

   Relay1 = 1
   Relay2 = 0
   Count1 = Count1 + 1
   Ii = Count1

   'On Urxc Baca_serial
   'Jika ada data serial masuk, baca sub baca_serial
   Enable Urxc
   Enable Interrupts

   Wait 1
   Next

Elseif Ii > 0 Then

   For I = 1 To 60                                          'Mengaktifkan VH 5 volt selama 60 detik
   Count2 = 0
   Yy = Getadc(0)
   Y1 = Yy * 5
   Y2 = Y1 / 1023

   'Rs= ((Vc/VRL)-1)*RL
   Rs = 5 / Y2                                              'PPM terbaca sekarang
   Rs = Rs - 1
   Rs = Rs * 10

   Teg = Rs / Ro

   'PPM = (30.34/ tegangan) ^ (100/53)
   X1 = 30.34 / Teg
   X2 = 100 / 53
   Ubah_ppm = X1 ^ X2


   Ubah_teg = Fusing(y2 , "#.##")
   Ppm = Int(ubah_ppm)

   Relay1 = 1
   Relay2 = 0
   Count1 = Count1 + 1
   Ii = Count1

   'On Urxc Baca_serial
   'Jika ada data serial masuk, baca sub baca_serial
   Enable Urxc
   Enable Interrupts

   Wait 1
   Next
End If

If Ii >= 60 Then
Ii = 1
   For J = 1 To 90                                          'Mengaktifkan VH 1.5 volt selama 90 detik
   Count1 = 0
   Yy = Getadc(0)
   Y1 = Yy * 5
   Y2 = Y1 / 1023

   'Rs= ((Vc/VRL)-1)*RL
   Rs = 5 / Y2                                              'PPM terbaca sekarang
   Rs = Rs - 1
   Rs = Rs * 10

   Teg = Rs / Ro

   'PPM = (30.34/ tegangan) ^ (100/53)
   X1 = 30.34 / Teg
   X2 = 100 / 53
   Ubah_ppm = X1 ^ X2


   Ubah_teg = Fusing(y2 , "#.##")
   Ppm = Int(ubah_ppm)

   Relay1 = 0
   Relay2 = 1
   Count2 = Count2 + 1
   Jj = Count2

   'On Urxc Baca_serial
   'Jika ada data serial masuk, baca sub baca_serial
   Enable Urxc
   Enable Interrupts
   Next
   Wait 1

Elseif Jj > 1 Then
   For J = Jj To 90                                         'Mengaktifkan VH 1.5 volt selama 90 detik
   Count1 = 0
   Yy = Getadc(0)
   Y1 = Yy * 5
   Y2 = Y1 / 1023

   'Rs= ((Vc/VRL)-1)*RL
   Rs = 5 / Y2                                              'PPM terbaca sekarang
   Rs = Rs - 1
   Rs = Rs * 10

   Teg = Rs / Ro

   'PPM = (30.34/ tegangan) ^ (100/53)
   X1 = 30.34 / Teg
   X2 = 100 / 53
   Ubah_ppm = X1 ^ X2


   Ubah_teg = Fusing(y2 , "#.##")
   Ppm = Int(ubah_ppm)

   Relay1 = 0
   Relay2 = 1
   Count2 = Count2 + 1
   Jj = Count2


   'On Urxc Baca_serial
   'Jika ada data serial masuk, baca sub baca_serial
   Enable Urxc
   Enable Interrupts
   Next
   Wait 1
End If

If Jj >= 90 Then
Jj = 1
Cls
Cursor Off
Locate 1 , 1
Lcd "Tegangan= " ; Ubah_teg ; " V"
Locate 2 , 1
Lcd "PPM CO  = " ; Ppm
Print Ubah_teg
Print Ppm
Wait 1


  If Ppm <= Batas_min Then
  Hijau = 1
  Kuning = 0
  Merah = 0
  Elseif Ppm > Batas_min And Ppm < Batas_max Then
  Hijau = 0
  Kuning = 1
  Merah = 0
  Elseif Ppm >= Batas_max Then
  Hijau = 0
  Kuning = 0
  Merah = 1
  End If
End If
Loop



Baca_serial:
'A: atas
'B: bawah
'C: kiri
'D: kanan
'E: stop
Disable Urxc
Disable Interrupts

Datas = Inkey()                                             'Membaca data serial yg masuk
Input Datas Noecho                                          'Data serial dibaca satu persatu
Pisah = Split(datas , Cek(1) , " ")                         'Data serial yg diterima dipisahkan berdasarkan spasi

If Datas = "A" Then
Ena = 1
In1 = 1
In2 = 0
Enb = 1
In3 = 1
In4 = 0
Elseif Datas = "B" Then
Ena = 1
In1 = 0
In2 = 1
Enb = 1
In3 = 0
In4 = 1
Elseif Datas = "C" Then
Ena = 1
In1 = 1
In2 = 0
Enb = 1
In3 = 0
In4 = 1
Elseif Datas = "D" Then
Ena = 1
In1 = 0
In2 = 1
Enb = 1
In3 = 1
In4 = 0
Elseif Datas = "E" Then
Ena = 1
In1 = 0
In2 = 0
Enb = 1
In3 = 0
In4 = 0
Elseif Cek(1) = "Y" Then                                    'data serial setelah spasi pertama
Batas_min = Val(cek(2))                                     'data serial setelah spasi kedua
Simpan_min = Batas_min                                      'simpan eeprom
Batas_max = Val(cek(3))                                     'data serial setelah spasi ketiga
Simpan_max = Batas_max                                      'simpan eeprom
Cls
Cursor Off
Locate 1 , 1
Lcd "Batas min= " ; Batas_min
Locate 2 , 1
Lcd "Batas max= " ; Batas_max
Wait 1
End If
Return