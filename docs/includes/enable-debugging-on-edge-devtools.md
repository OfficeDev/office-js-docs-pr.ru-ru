Если надстройка работает в Microsoft Edge, код без пользовательского интерфейса не сможет по умолчанию подключаться к отладчику.
Код без пользовательского интерфейса — это любой код, выполняющийся, когда область задач не отображается, например команды надстройки. Чтобы включить отладку, требуется выполнить следующую команду [Windows PowerShell](https://docs.microsoft.com/powershell/scripting/getting-started/getting-started-with-windows-powershell):

1. Выполните следующую команду, чтобы получить сведения о пакете приложения **Microsoft.Win32WebViewHost**.
    
    ```powershell
    Get-AppxPackage Microsoft.Win32WebViewHost
    ```
    
    Эта команда перечисляет сведения о пакете приложения аналогично следующему результату.
    
    ```powershell
    Name              : Microsoft.Win32WebViewHost
    Publisher         : CN=Microsoft Windows, O=Microsoft Corporation, L=Redmond, S=Washington, C=US
    Architecture      : Neutral
    ResourceId        : neutral
    Version           : 10.0.18362.449
    PackageFullName   : Microsoft.Win32WebViewHost_10.0.18362.449_neutral_neutral_cw5n1h2txyewy
    InstallLocation   : C:\Windows\SystemApps\Microsoft.Win32WebViewHost_cw5n1h2txyewy
    IsFramework       : False
    PackageFamilyName : Microsoft.Win32WebViewHost_cw5n1h2txyewy
    PublisherId       : cw5n1h2txyewy
    IsResourcePackage : False
    IsBundle          : False
    IsDevelopmentMode : False
    NonRemovable      : True
    IsPartiallyStaged : False
    SignatureKind     : System
    Status            : Ok
    ```
    
2. Чтобы включить отладку, выполните следующую команду. Используйте значение для **PackageFullName**, полученное в списке от предыдущей команды.
    
    ```powershell
    setx JS_DEBUG <PackageFullName>
    ```
    
3. Если Office уже запущен, закройте и перезапустите его, чтобы учесть изменения отладки.