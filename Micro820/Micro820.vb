
Imports System.Threading.Tasks
Imports EasyModbus
Public Class Micro820
    Dim modbus As New ModbusClient
    Const Count As Int16 = 229
    Dim Coils(Count) As Boolean
    Dim _IP As String
    Dim sw As New Stopwatch
    Public Event Cycle(arg As EventArgss)
    Dim args As New EventArgss
    Dim sw2 As New Stopwatch


    Sub New(IP As String)
        modbus.Port = 502
        modbus.IPAddress = IP
        _IP = IP
    End Sub
    Public Sub Connect()
        Try
            If Not modbus.Connected Then modbus.Connect()
            Console.WriteLine("Micro 820 is Connected")
            Task.Factory.StartNew(Sub() ThreadStart())
        Catch ex As Exception

            Throw New Exception(String.Format("IP:{0} Cihaza Bağlanılamadı!", _IP))
        End Try


    End Sub
    Public ReadOnly Property isConnected() As Boolean
        Get
            Return modbus.Connected
        End Get
    End Property

    Public Class EventArgss
        Property CycleTime As Integer
        Property Hız As Single
        Property İniş As Single
        Property Kalkış As Single
        Property Ready As Boolean
        Property Run As Boolean
        Property Errors As Boolean
        Property ErrorCode As Int16
        Property ClrFault As Boolean
        Property Sayac As Integer
        Property ShrinkSayısıTimer As Integer
        Property ShrinkYarısıTimer As Integer
        Property ShrinkSayısı As Integer
        Property ShrinkPistonu As Boolean
        Property ShrinkBol As Integer




    End Class
    Private Enum Adresler
        Ready = 0               '1bit uzunluğunda bool
        Run = 1                 '1bit uzunluğunda bool
        ErrorS = 2              '1bit uzunluğunda bool
        ErrorCode = 3           '16bit uzunluğunda integer
        ClrFault = 19           '1bit uzunluğunda bool
        Sayac = 20              '16bit uzunluğunda integer
        ShrinkyarısıTimer = 36  '16bit uzunluğunda integer
        ShrinkSayısıTimer = 52  '16bit uzunluğunda integer
        ShrinkSayısı = 68       '16bit uzunluğunda integer
        ShrinkPistonu = 84      '1bit uzunluğunda bool
        SpeedYaz = 85           '32bit uzunluğunda single
        SpeedOku = 117          '32bit uzunluğunda single
        İniş = 149              '32bit uzunluğunda single
        Kalkış = 181            '32bit uzunluğunda single
        ShrinkBol = 213         '16bit uzunluğunda integer
    End Enum

    'Dim gönderbin As String = StrReverse(ToBin(value, BitTipi.bit_16))
    'Dim gönderbit = gönderbin.ToCharArray.ToBoolean
    ' Modbus.WriteMultipleCoils(Adresler.DışTepsiHızı, gönderbit)
    Property Hız As Single
        Get
            Dim Speedarry = Coils.SubArray(Adresler.SpeedOku, 32)
            Dim Speedstr = String.Join("", Speedarry).Replace("True", 1).Replace("False", 0)
            Return BinaryStringToSingle(StrReverse(Speedstr))
        End Get
        Set(value As Single)
            Dim gönderbin As String = StrReverse(SingleToBinaryString(value))
            Dim gönderbit = gönderbin.ToCharArray.ToBoolean
            modbus.WriteMultipleCoils(Adresler.SpeedYaz, gönderbit)
        End Set
    End Property
    Property İniş As Single
        Get
            Dim Speedarry = Coils.SubArray(Adresler.İniş, 32)
            Dim Speedstr = String.Join("", Speedarry).Replace("True", 1).Replace("False", 0)
            Return BinaryStringToSingle(StrReverse(Speedstr))
        End Get
        Set(value As Single)
            Dim gönderbin As String = StrReverse(SingleToBinaryString(value))
            Dim gönderbit = gönderbin.ToCharArray.ToBoolean
            modbus.WriteMultipleCoils(Adresler.İniş, gönderbit)
        End Set
    End Property
    Property Kalkış As Single
        Get
            Dim Speedarry = Coils.SubArray(Adresler.Kalkış, 32)
            Dim Speedstr = String.Join("", Speedarry).Replace("True", 1).Replace("False", 0)
            Return BinaryStringToSingle(StrReverse(Speedstr))
        End Get
        Set(value As Single)
            Dim gönderbin As String = StrReverse(SingleToBinaryString(value))
            Dim gönderbit = gönderbin.ToCharArray.ToBoolean
            modbus.WriteMultipleCoils(Adresler.Kalkış, gönderbit)
        End Set
    End Property
    ReadOnly Property Ready As Boolean
        Get
            Return Coils(Adresler.Ready)
        End Get
    End Property
    ReadOnly Property Run As Boolean
        Get
            Return Coils(Adresler.Run)
        End Get
    End Property
    ReadOnly Property Errors As Boolean
        Get
            Return Coils(Adresler.ErrorS)
        End Get
    End Property
    ReadOnly Property ErrorCode As Integer
        Get
            Dim Speedarry = Coils.SubArray(Adresler.ErrorCode, 16)
            Dim Speedstr = String.Join("", Speedarry).Replace("True", 1).Replace("False", 0)
            Return Convert.ToInt32(StrReverse(Speedstr), 2)
        End Get
    End Property
    Property ClrFault As Boolean
        Get
            Return Coils(Adresler.ClrFault)
        End Get
        Set(value As Boolean)
            modbus.WriteSingleCoil(Adresler.ClrFault, value)
        End Set
    End Property
    Property Sayac As Integer
        Get
            Dim Speedarry = Coils.SubArray(Adresler.Sayac, 16)
            Dim Speedstr = String.Join("", Speedarry).Replace("True", 1).Replace("False", 0)
            Return Convert.ToInt32(StrReverse(Speedstr), 2)
        End Get
        Set(value As Integer)
            Dim gönderbin As String = StrReverse(ToBin(value, BitTipi.bit_16))
            Dim gönderbit = gönderbin.ToCharArray.ToBoolean
            modbus.WriteMultipleCoils(Adresler.Sayac, gönderbit)
        End Set
    End Property
    Property ShrinkBol As Integer
        Get
            Dim Speedarry = Coils.SubArray(Adresler.ShrinkBol, 16)
            Dim Speedstr = String.Join("", Speedarry).Replace("True", 1).Replace("False", 0)
            Return Convert.ToInt32(StrReverse(Speedstr), 2)
        End Get
        Set(value As Integer)
            Dim gönderbin As String = StrReverse(ToBin(value, BitTipi.bit_16))
            Dim gönderbit = gönderbin.ToCharArray.ToBoolean
            modbus.WriteMultipleCoils(Adresler.ShrinkBol, gönderbit)
        End Set
    End Property
    Property ShrinkSayısıTimer As Integer
        Get
            Dim Speedarry = Coils.SubArray(Adresler.ShrinkSayısıTimer, 16)
            Dim Speedstr = String.Join("", Speedarry).Replace("True", 1).Replace("False", 0)
            Return Convert.ToInt32(StrReverse(Speedstr), 2)
        End Get
        Set(value As Integer)
            Dim gönderbin As String = StrReverse(ToBin(value, BitTipi.bit_16))
            Dim gönderbit = gönderbin.ToCharArray.ToBoolean
            modbus.WriteMultipleCoils(Adresler.ShrinkSayısıTimer, gönderbit)
        End Set
    End Property
    Property ShrinkYarısıTimer As Integer
        Get
            Dim Speedarry = Coils.SubArray(Adresler.ShrinkyarısıTimer, 16)
            Dim Speedstr = String.Join("", Speedarry).Replace("True", 1).Replace("False", 0)
            Return Convert.ToInt32(StrReverse(Speedstr), 2)
        End Get
        Set(value As Integer)
            Dim gönderbin As String = StrReverse(ToBin(value, BitTipi.bit_16))
            Dim gönderbit = gönderbin.ToCharArray.ToBoolean
            modbus.WriteMultipleCoils(Adresler.ShrinkyarısıTimer, gönderbit)
        End Set
    End Property
    Property ShrinkSayısı As Integer
        Get
            Dim Speedarry = Coils.SubArray(Adresler.ShrinkSayısı, 16)
            Dim Speedstr = String.Join("", Speedarry).Replace("True", 1).Replace("False", 0)
            Return Convert.ToInt32(StrReverse(Speedstr), 2)
        End Get
        Set(value As Integer)
            Dim gönderbin As String = StrReverse(ToBin(value, BitTipi.bit_16))
            Dim gönderbit = gönderbin.ToCharArray.ToBoolean
            modbus.WriteMultipleCoils(Adresler.ShrinkSayısı, gönderbit)
        End Set
    End Property
    ReadOnly Property ShrinkPistonu As Boolean
        Get
            Return Coils(Adresler.ShrinkPistonu)
        End Get
    End Property
    Sub ThreadStart()
        Try
            Do Until Not modbus.Connected
                sw.Start()
                Try
                    Coils = modbus.ReadCoils(0, Count)
                Catch ex As Exception
                    Console.WriteLine(ex.ToString)
                End Try

                sw.Stop()
                'Console.WriteLine("Thread Time:{0}ms", sw.ElapsedMilliseconds)
                ''------------
                Task.Factory.StartNew(Sub() EventStart(sw.ElapsedMilliseconds))

                sw.Reset()
            Loop
            modbus.Disconnect()
            Throw New Exception(String.Format("IP:{0} Cihaz Bağlantısı Kesildi!", _IP))
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try


    End Sub
    Sub EventStart(cycletime As Integer)
        sw2.Start()
        'Console.WriteLine("Cycling Start...")
        With args
            .Ready = Ready
            .Hız = Hız
            .ClrFault = ClrFault
            .CycleTime = cycletime
            .ErrorCode = ErrorCode
            .Errors = Errors
            .Kalkış = Kalkış
            .Run = Run
            .Sayac = Sayac
            .ShrinkPistonu = ShrinkPistonu
            .ShrinkSayısı = ShrinkSayısı
            .ShrinkSayısıTimer = ShrinkSayısıTimer
            .ShrinkYarısıTimer = ShrinkYarısıTimer
            .İniş = İniş
            .ShrinkBol = ShrinkBol
        End With
        sw2.Stop()
        args.CycleTime = sw2.ElapsedMilliseconds
        sw2.Reset()
        RaiseEvent Cycle(args)
    End Sub
    Private Enum BitTipi
        bit_16 = 16
        bit_32 = 32
    End Enum
    Private Function ToBin(ByVal value As Integer, ByVal len As BitTipi) As String
        Return (If(len > 1, ToBin(value >> 1, len - 1), Nothing)) & "01".Chars(value And 1)
    End Function
    ''KUllanılmıyıcak
    Private Function DoubleToBinaryString(ByVal d As Double) As String
        Return Convert.ToString(BitConverter.DoubleToInt64Bits(d), 2)
    End Function
    ''KUllanılmıyıcak
    Private Function BinaryStringToDouble(ByVal s As String) As Double
        Return BitConverter.Int64BitsToDouble(Convert.ToInt64(s, 2))
    End Function
    Private Function SingleToBinaryString(ByVal f As Single) As String
        Dim b() As Byte = BitConverter.GetBytes(f)
        Dim i As Integer = BitConverter.ToInt32(b, 0)
        Return Convert.ToString(i, 2)
    End Function

    Private Function BinaryStringToSingle(ByVal s As String) As Single
        Dim i As Integer = Convert.ToInt32(s, 2)
        Dim b() As Byte = BitConverter.GetBytes(i)
        Return BitConverter.ToSingle(b, 0)
    End Function
    Private Function IntegerToBinaryString(ByVal d As Integer) As String
        Return d.ToString(2)
    End Function
    Private Function BinaryStringToInteger(d As String) As Integer
        Return Convert.ToInt32(d, 2)
    End Function
End Class
Public Module Extensions
    <System.Runtime.CompilerServices.Extension>
    Public Function SubArray(Of T)(ByVal data() As T, ByVal index As Integer, ByVal length As Integer) As T()
        Dim result(length - 1) As T
        Array.Copy(data, index, result, 0, length)
        Return result
    End Function
    <System.Runtime.CompilerServices.Extension>
    Public Function ToBoolean(ByVal Gelen As Char()) As Boolean()
        Dim dönen(Gelen.Length - 1) As Boolean
        Dim index As Integer = 0
        For Each item In Gelen
            dönen(index) = If(item = "1"c, True, False)
            index += 1
        Next
        Return dönen
    End Function



End Module