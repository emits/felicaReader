
Imports System.Runtime.InteropServices
Imports UInt8 = System.Byte

Module modFelicaLib
    Private Const MAX_SYSTEM_CODE = 8
    Private Const MAX_AREA_CODE = 16
    Private Const MAX_SERVICE_CODE = 256

    '-------------------------------
    ' FeliCa の情報を格納する構造体
    '-------------------------------
    <StructLayout(LayoutKind.Sequential)> _
    Private Structure felica
        Public p As IntPtr           'PaSoRi ハンドル
        Public systemcode As UInt16  'システムコード
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)> _
        Public IDm() As UInt8        'IDm
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)> _
        Public PMm() As UInt8        'PMm

        'systemcode
        Public num_system_code As UInt8      '列挙システムコード数
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=MAX_SYSTEM_CODE)> _
        Public system_code() As UInt16       '列挙システムコード

        'area/service codes
        Public num_area_code As UInt8        'エリアコード数
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=MAX_AREA_CODE)> _
        Public area_code() As UInt16         'エリアコード
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=MAX_AREA_CODE)> _
        Public end_service_code() As UInt16  'エンドサービスコード

        Public num_service_code As UInt8     'サービスコード数
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=MAX_SERVICE_CODE)> _
        Public service_code() As UInt16      'サービスコード
    End Structure

    '-----------
    ' constants
    '-----------
    'システムコード (ネットワークバイトオーダ/ビックエンディアンで表記)
    Private Const POLLING_ANY = &HFFFF
    Private Const POLLING_EDY = &HFE00  'システムコード: 共通領域 (Edyなどが使用)
    Private Const POLLING_SUICA = &H3   'システムコード: サイバネ領域

    '------
    ' APIs
    '------
    '【関 数 名】pasori_open
    '【第１引数】[in] dummy
    '【戻 り 値】pasori  ハンドル
    <DllImport("felicalib.dll")> _
    Private Function pasori_open(ByVal dummy As IntPtr) As IntPtr
    End Function

    '【関 数 名】pasori_close
    '【第１引数】[in] p  pasoriハンドル (pasori_open で得たポインタを指定する)
    <DllImport("felicalib.dll")> _
    Private Sub pasori_close(ByVal p As IntPtr)
    End Sub

    '【関 数 名】InteropServices
    '【第１引数】[in] p  pasoriハンドル (pasori_open で得たポインタを指定する)
    '【戻 り 値】エラーコード
    <DllImport("felicalib.dll")> _
    Private Function pasori_init(ByVal p As IntPtr) As Integer
    End Function

    '【関 数 名】felica_polling
    '【第１引数】[in] p           pasoriハンドル (pasori_open で得たポインタを指定する)
    '【第２引数】[in] systemcode  システムコード
    '【第３引数】[in] RFU         RFU (使用しない)
    '【第４引数】[in] timeslot    タイムスロット
    '【戻 り 値】felicaハンドル (MarshalクラスのPtrToStructureメソッドを使用)
    <DllImport("felicalib.dll")> _
    Private Function felica_polling( _
        ByVal p As IntPtr, _
        ByVal systemcode As UInt16, _
        ByVal RFU As UInt8, _
        ByVal timeslot As UInt8 _
    ) As IntPtr
    End Function

    '【関 数 名】felica_free
    '【第１引数】[in] f  felicaハンドル (felica構造体のポインタを指定する)
    <DllImport("felicalib.dll")> _
    Private Sub felica_free(ByVal f As IntPtr)
    End Sub

    '【関 数 名】felica_getidm
    '【第１引数】[in]  f   felicaハンドル (felica構造体のポインタを指定する)
    '【第２引数】[out] buf IDm を格納するバッファ(8バイト)
    <DllImport("felicalib.dll")> _
    Private Sub felica_getidm(ByVal f As IntPtr, ByVal buf As IntPtr)
    End Sub

    '【関 数 名】felica_getpmm
    '【第１引数】[in]   f    felicaハンドル (felica構造体のポインタを指定する)
    '【第２引数】[out]  buf  PMm を格納するバッファ(8バイト)
    <DllImport("felicalib.dll")> _
    Private Sub felica_getpmm(ByVal f As IntPtr, ByVal buf As IntPtr)
    End Sub

    '【関 数 名】felica_read_without_encryption02
    '【第１引数】[in]   f            felicaハンドル (felica構造体のポインタを指定する)
    '【第２引数】[in]   servicecode  サービスコード
    '【第３引数】[in]   mode         モード(使用しない)
    '【第４引数】[in]   addr         ブロック番号
    '【第５引数】[out]  data         データ(16バイト)
    '【戻 り 値】エラーコード
    <DllImport("felicalib.dll")> _
    Private Function felica_read_without_encryption02( _
        ByVal f As IntPtr, _
        ByVal servicecode As Integer, _
        ByVal mode As Integer, _
        ByVal addr As UInt8, _
        ByVal buf As IntPtr _
    ) As Integer
    End Function

    '【関 数 名】felica_write_without_encryption
    '【第１引数】[in]   f            felicaハンドル (felica構造体のポインタを指定する)
    '【第２引数】[in]   servicecode  サービスコード
    '【第３引数】[in]   mode         モード(使用しない)
    '【第４引数】[in]   addr         ブロック番号
    '【第５引数】[out]  data         データ(16バイト)
    '【戻 り 値】エラーコード
    <DllImport("felicalib.dll")> _
    Private Function felica_write_without_encryption( _
        ByVal f As IntPtr, _
        ByVal servicecode As Integer, _
        ByVal addr As UInt8, _
        ByVal buf As IntPtr _
    ) As Integer
    End Function

    '【関 数 名】felica_enum_systemcode
    '【第１引数】[in]  p  pasoriハンドル (pasori_open で得たポインタを指定する)
    '【戻 り 値】felicaハンドル (felica構造体のポインタを指定する)
    <DllImport("felicalib.dll")> _
    Private Function felica_enum_systemcode(ByVal p As IntPtr) As IntPtr
    End Function

    '【関 数 名】felica_enum_service
    '【第１引数】[in]  p           pasoriハンドル (pasori_open で得たポインタを指定する)
    '【第２引数】[in]  systemcode  システムコード
    '【戻 り 値】felicaハンドル (felica構造体のポインタを指定する)
    <DllImport("felicalib.dll")> _
    Private Function felica_enum_service(ByVal p As IntPtr, ByVal systemcode As UInt16) As IntPtr
    End Function


    '===================================
    ' Win32 API (DLL存在チェックに必要)
    '===================================
    <DllImport("kernel32")> _
    Private Function SetDllDirectory(ByVal lpPathName As String) As Boolean
    End Function

    <DllImport("kernel32")> _
    Private Function LoadLibrary(ByVal lpLibFileName As String) As Integer
    End Function

    <DllImport("kernel32")> _
    Private Function GetProcAddress(ByVal hModule As Integer, ByVal lpProcName As String) As Integer
    End Function

    <DllImport("kernel32")> _
    Private Function FreeLibrary(ByVal hLibModule As Integer) As Boolean
    End Function

    '====================================================
    '【関数名】isDLLExists
    '【引  数】なし
    '【戻り値】[out] Boolean  TRUE   DLLの読み込みに成功
    '                         FALSE  DLLの読み込みに失敗
    '----------------------------------------------------
    ' felicalib.dll が存在するかチェックする。
    '====================================================
    Public Function isDLLExists() As Boolean
        'DLLのハンドル
        Dim hModule As IntPtr = IntPtr.Zero

        'DLL内の関数を受け取るポインタ
        Dim pPasoriOpen As IntPtr = IntPtr.Zero
        Dim pPasoriClose As IntPtr = IntPtr.Zero
        Dim pPasoriInit As IntPtr = IntPtr.Zero
        Dim pFelicaPolling As IntPtr = IntPtr.Zero
        Dim pFelicaFree As IntPtr = IntPtr.Zero
        Dim pFelicaGetidm As IntPtr = IntPtr.Zero
        Dim pFelicaGetpmm As IntPtr = IntPtr.Zero
        Dim pFelicaReadWithoutEncryption02 As IntPtr = IntPtr.Zero
        Dim pFelicaWriteWithoutEncryption As IntPtr = IntPtr.Zero
        Dim pFelicaEnumSystemcode As IntPtr = IntPtr.Zero
        Dim pFelicaEnumService As IntPtr = IntPtr.Zero

        Try
            'DLLプリロード対策 ※DLLの読み込み先から現在の作業ディレクトリ(CWD)を除外する
            SetDllDirectory("")

            'DLLを読み込む
            hModule = LoadLibrary("felicalib.dll")
            If hModule = IntPtr.Zero Then
                Return False
            End If

            'DLLの関数を読み込む
            pPasoriOpen = GetProcAddress(hModule, "pasori_open")
            pPasoriClose = GetProcAddress(hModule, "pasori_close")
            pPasoriInit = GetProcAddress(hModule, "pasori_init")
            pFelicaPolling = GetProcAddress(hModule, "felica_polling")
            pFelicaFree = GetProcAddress(hModule, "felica_free")
            pFelicaGetidm = GetProcAddress(hModule, "felica_getidm")
            pFelicaGetpmm = GetProcAddress(hModule, "felica_getpmm")
            pFelicaReadWithoutEncryption02 = GetProcAddress(hModule, "felica_read_without_encryption02")
            pFelicaWriteWithoutEncryption = GetProcAddress(hModule, "felica_write_without_encryption")
            pFelicaEnumSystemcode = GetProcAddress(hModule, "felica_enum_systemcode")
            pFelicaEnumService = GetProcAddress(hModule, "felica_enum_service")

            If pPasoriOpen = IntPtr.Zero OrElse _
               pPasoriClose = IntPtr.Zero OrElse _
               pPasoriInit = IntPtr.Zero OrElse _
               pFelicaPolling = IntPtr.Zero OrElse _
               pFelicaFree = IntPtr.Zero OrElse _
               pFelicaGetidm = IntPtr.Zero OrElse _
               pFelicaGetpmm = IntPtr.Zero OrElse _
               pFelicaReadWithoutEncryption02 = IntPtr.Zero OrElse _
               pFelicaWriteWithoutEncryption = IntPtr.Zero OrElse _
               pFelicaEnumSystemcode = IntPtr.Zero OrElse _
               pFelicaEnumService = IntPtr.Zero _
            Then
                FreeLibrary(hModule)
                Return False
            End If

            '読み込み成功
            FreeLibrary(hModule)
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "isDLLExists()")
            If hModule <> IntPtr.Zero Then FreeLibrary(hModule)
            Return False
        End Try
    End Function

    '=============================================
    '【関 数 名】hexdump
    '【第１引数】[in]  UInt8()  データ配列
    '【第２引数】[in]  Integer  配列のサイズ
    '【戻 り 値】[out] String   16進数文字列
    '---------------------------------------------
    ' 受け取ったデータを16進数の文字列に変換する。
    '=============================================
    Private Function hexdump(ByVal arg() As UInt8, ByVal size As Integer) As String
        Dim sResult As String = ""

        For I As Integer = 0 To size - 1
            sResult &= arg(I).ToString("X2")
        Next

        Return sResult
    End Function

    '================================
    ' felicalib.dll アクセス用クラス
    '================================
    Class CFelicaLib
        Implements IDisposable 'デストラクタ

        '==========
        ' 変数定義
        '==========
        Private p_ptr As IntPtr  'Pasoriポインタ
        Private f_ptr As IntPtr  'felica構造体ポインタ

        '===========================================
        '【関数名】New
        '【引  数】なし
        '【戻り値】なし
        '-------------------------------------------
        ' コンストラクタ。ポインタの初期化を行なう。
        '===========================================
        Public Sub New()
            '初期化
            p_ptr = IntPtr.Zero
            f_ptr = IntPtr.Zero
        End Sub

        '=================================
        '【関数名】Dispose
        '【引  数】なし
        '【戻り値】なし
        '---------------------------------
        ' デストラクタ。PaSoRiを解放する。
        '=================================
        Public Sub Dispose() Implements IDisposable.Dispose
            'PaSoRi の接続を解放
            Pasori_Free()
        End Sub

        '===============================================
        '【関数名】Pasori_Connect
        '【引  数】なし
        '【戻り値】[out] Boolean  TRUE   PaSoRi接続成功
        '                         FALSE  PaSoRi接続失敗
        '-----------------------------------------------
        ' PaSoRiに接続して使用可能な状態にする。
        '===============================================
        Public Function Pasori_Connect() As Boolean
            Try
                '----------------------
                ' PaSoRi ハンドル取得
                '----------------------
                p_ptr = pasori_open(Nothing)
                If p_ptr = IntPtr.Zero Then
                    Return False
                End If

                '---------------------
                ' PaSoRi 初期化(接続)
                '---------------------
                If pasori_init(p_ptr) <> 0 Then
                    Return False
                End If

                Return True
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Pasori_Connect()")
                Return False
            End Try
        End Function

        '======================
        '【関数名】Pasori_Free
        '【引  数】なし
        '【戻り値】なし
        '----------------------
        ' PaSoRiの接続を解放
        '======================
        Public Sub Pasori_Free()
            Try
                If f_ptr <> IntPtr.Zero Then felica_free(f_ptr)
                If p_ptr <> IntPtr.Zero Then pasori_close(p_ptr)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Pasori_Free()")
            End Try
        End Sub

        '===========================================================
        '【関数名】Polling
        '【引  数】なし
        '【戻り値】[out] Boolean  TRUE   felicaハンドルの取得に成功
        '                         FALSE  felicaハンドルの取得に失敗
        '-----------------------------------------------------------
        ' ポーリング。FeliCaの読み取り準備。
        '===========================================================
        Public Function Polling() As Boolean
            Try
                '--------------------------------
                ' felicaハンドルを一度クリアする
                '--------------------------------
                If f_ptr <> IntPtr.Zero Then
                    felica_free(f_ptr)
                End If

                '------------
                ' ポーリング
                '------------
                f_ptr = felica_polling(p_ptr, POLLING_ANY, 0, 0)

                '------------
                ' 結果を返す
                '------------
                If f_ptr = IntPtr.Zero Then
                    Return False
                Else
                    Return True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Polling()")
                Return False
            End Try
        End Function

        '====================================
        '【関数名】getIDm
        '【引  数】なし
        '【戻り値】[out] String  FeliCaのIDm
        '------------------------------------
        ' FeliCaのIDmを取得する。
        '====================================
        Public Function getIDm() As String
            Try
                '------------
                ' エラー処理
                '------------
                If f_ptr = IntPtr.Zero Then
                    Return ""
                End If

                '-------------
                ' IDm読み取り
                '-------------
                Dim IDm As String
                Dim buf(8) As UInt8

                '---------------------
                ' bufのアドレスを取得
                '---------------------
                Dim gch As GCHandle = GCHandle.Alloc(buf, GCHandleType.Pinned)
                Dim b As IntPtr = gch.AddrOfPinnedObject().ToInt32

                felica_getidm(f_ptr, b)
                IDm = hexdump(buf, 8)
                gch.Free()

                Return IDm
            Catch ex As Exception
                MessageBox.Show(ex.Message, "getIDm()")
                Return ""
            End Try
        End Function

        '====================================
        '【関数名】getPMm
        '【引  数】なし
        '【戻り値】[out] String  FeliCaのPMm
        '------------------------------------
        ' FeliCaのPMmを取得する。
        '====================================
        Public Function getPMm() As String
            Try
                '------------
                ' エラー処理
                '------------
                If f_ptr = IntPtr.Zero Then
                    Return ""
                End If

                '-------------
                ' PMm読み取り
                '-------------
                Dim PMm As String
                Dim buf(8) As UInt8

                '---------------------
                ' bufのアドレスを取得
                '---------------------
                Dim gch As GCHandle = GCHandle.Alloc(buf, GCHandleType.Pinned)
                Dim b As IntPtr = gch.AddrOfPinnedObject().ToInt32

                felica_getpmm(f_ptr, b)
                PMm = hexdump(buf, 8)
                gch.Free()

                Return PMm
            Catch ex As Exception
                MessageBox.Show(ex.Message, "getIDm()")
                Return ""
            End Try
        End Function

        '======================================
        '【関 数 名】getIDmPMm
        '【第１引数】[out] String  FeliCaのIDm
        '【第２引数】[out] String  FeliCaのPMm
        '【戻 り 値】なし
        '--------------------------------------
        ' FeliCaのIDmとPMmを同時に取得する。
        '======================================
        Public Sub getIDmPMm(ByRef IDm As String, ByRef PMm As String)
            Try
                '------------
                ' エラー処理
                '------------
                If f_ptr = IntPtr.Zero Then
                    IDm = ""
                    PMm = ""
                    Return
                End If

                '--------------------------
                ' felica構造体の実体を取得
                '--------------------------
                Dim f As felica = Marshal.PtrToStructure(f_ptr, GetType(felica))

                '------------------
                ' IDm・PMm読み取り
                '------------------
                IDm = hexdump(f.IDm, 8)
                PMm = hexdump(f.PMm, 8)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "getIDmPMm()")
                IDm = ""
                PMm = ""
            End Try
        End Sub
    End Class
End Module