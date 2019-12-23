<span data-ttu-id="b8c8c-101">Если надстройка работает в Microsoft Edge, код без пользовательского интерфейса не сможет по умолчанию подключаться к отладчику.</span><span class="sxs-lookup"><span data-stu-id="b8c8c-101">When the add-in is running in Microsoft Edge, UI-less code will not be able to attach to a debugger by default.</span></span>
<span data-ttu-id="b8c8c-102">Код без пользовательского интерфейса — это любой код, выполняющийся, когда область задач не отображается, например команды надстройки.</span><span class="sxs-lookup"><span data-stu-id="b8c8c-102">UI-less code is any code running while the task pane is not visible, such as add-in commands.</span></span> <span data-ttu-id="b8c8c-103">Чтобы включить отладку, требуется выполнить следующую команду [Windows PowerShell](https://docs.microsoft.com/powershell/scripting/getting-started/getting-started-with-windows-powershell):</span><span class="sxs-lookup"><span data-stu-id="b8c8c-103">To enable debugging, you need to run the following [Windows PowerShell](https://docs.microsoft.com/powershell/scripting/getting-started/getting-started-with-windows-powershell) commands.</span></span>

1. <span data-ttu-id="b8c8c-104">Выполните следующую команду, чтобы получить сведения о пакете приложения **Microsoft.Win32WebViewHost**.</span><span class="sxs-lookup"><span data-stu-id="b8c8c-104">Run the following command to get information for the **Microsoft.Win32WebViewHost** app package.</span></span>
    
    ```powershell
    Get-AppxPackage Microsoft.Win32WebViewHost
    ```
    
    <span data-ttu-id="b8c8c-105">Эта команда перечисляет сведения о пакете приложения аналогично следующему результату.</span><span class="sxs-lookup"><span data-stu-id="b8c8c-105">The command lists app package information similar to the following output.</span></span>
    
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
    
2. <span data-ttu-id="b8c8c-106">Чтобы включить отладку, выполните следующую команду:</span><span class="sxs-lookup"><span data-stu-id="b8c8c-106">Run the following command to verify content filtering is enabled:</span></span> <span data-ttu-id="b8c8c-107">Используйте значение для **PackageFullName**, полученное в списке от предыдущей команды.</span><span class="sxs-lookup"><span data-stu-id="b8c8c-107">Use the value for the **PackageFullName** listed from the previous command.</span></span>
    
    ```powershell
    setx JS_DEBUG <PackageFullName>
    ```
    
3. <span data-ttu-id="b8c8c-108">Если Office уже запущен, закройте и перезапустите его, чтобы учесть изменения отладки.</span><span class="sxs-lookup"><span data-stu-id="b8c8c-108">If Office was already running, close and restart Office so that it picks up the debugging change.</span></span>