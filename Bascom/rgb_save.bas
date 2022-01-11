'Pinout ATtiny25:
'1- PB5 (PCINT5 / Reset / ADC0 / DW)
'2- PB3 (PCINT3/XTAL1/CLKI/OC1B/ADC3)
'3- PB4 (PCINT4/XTAL2/CLKO/_OC1B/ADC2)
'4- GND
'5- PB0 (MOSI/DI/SDA/AIN0/OC0A/OC1A/AREF/PCINT0)
'6- PB1 (MISO/DO/AIN1/OC0B/OC1A/PCINT1)
'7- PB2 (SCK/USCK/SCL/ADC1/T0/INT0/PCINT2)
'8- VCC

'             PCINT5/RESET/ADC0/DW PB5 [1 U 8] VCC
' PCINT3/XTAL1/CLKI/OC1B/ADC3 PB3 [2    7] PB2 SCK/USCK/SCL/ADC1/T0/INT0/PCINT2
'PCINT4/XTAL2/CLKO/OC1B/ADC2 PB4 [3    6] PB1 MISO/DO/AIN1/OC0B/OC1A/PCINT1
'                                                   GND [4    5] PB0 MOSI/DI/SDA/AIN0/OC0A/OC1A/AREF/PCINT0
'
'
$regfile = "attiny25.dat"
$crystal = 16000000
'$hwstack = 64
'$swstack = 32
'$framesize = 32

'--------------------------------------Software UART-------------------------------------
Open "comb.3:4800,8,n,1" For Input As #1

Const Red = 0
Const Green = 1
Const Blue = 4

Dim R , G , B As Byte
Dim Value As Byte
Dim Ident As Byte
Dim Ee_r , Ee_g , Ee_b As Eram Byte

'---------------------------------------DISPLAY-------------------------------------------
Main:
    Config Portb.red = Output
    Config Portb.green = Output
    Config Portb.blue = Output

'Configure counter/timer0 for fast PWM on PB0 and PB1
 'COM0A1 COM0A0 COM0B1 COM0B0 – – WGM01 WGM00
    Tccr0a = &B11110011                                     '64khz
 ' FOC0A FOC0B – – WGM02 CS02 CS01 CS00
    Tccr0b = &B00000001                                     '64khz
'Configure counter/timer1 for fast PWM on PB4
' TSM PWM1B COM1B1 COM1B0 FOC1B FOC1A PSR1 PSR0
    Gtccr = &B01110000
'CTC1 PWM1A COM1A1 COM1A0 CS13 CS12 CS11 CS10
    Tccr1 = &B00110100                                      '32khz

'---------------------------PLL AS CLOCK SOURCE for 16MHz---------------------------
    Pllcsr.plle = 1
    While Pllcsr.plock = 0
    Wend
    Pllcsr.pcke = 1
    !Ldi R24 , 85                                           '16MHZ
    !Out Osccal , R24
'----------------------------------SHOW PWM VALUES-----------------------------------
    R = Ee_r
    G = Ee_g
    B = Ee_b

    Ocr0a = R
    Ocr0b = G
    Ocr1b = B

    'espera que a comunicação USB estabilize
    Wait 3
    Enable Interrupts

    'Saída 62500Hz exactos para baud correto
    'Ocr0a = 127
    'Ocr0b = 127
    'Ocr1b = 127
    'Do
    'Loop

    'O valor é recebido 2 vezes, 1ª identificação,  2ª valor a exibir/gravar
    Do
        Value = Waitkey(#1)
        Ident = Value And 3
        If Value < 4 Then
            Value = Waitkey(#1)
            Select Case Ident
                Case 0 : R = 255 - Value
                Case 1 : G = 255 - Value
                Case 2 : B = 255 - Value
                Case 3 :
                    Ee_r = R
                    Ee_g = G
                    Ee_b = B
                     'pequeno flash
                    Ocr0a = 0
                    Ocr0b = 0
                    Ocr1b = 0
                    Waitms 10
                    Ocr0a = R
                    Ocr0b = G
                    Ocr1b = B
            End Select
        End If
        Ocr0a = R
        Ocr0b = G
        Ocr1b = B
    Loop

    End