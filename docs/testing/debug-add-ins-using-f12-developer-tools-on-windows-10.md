---
title: Отладка надстроек с помощью средств разработчика F12 в Windows 10
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 57e0818da0a115380bcb11f1226c48d3ded1078a
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2018
ms.locfileid: "23943728"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a><span data-ttu-id="1e268-102">Отладка надстроек с помощью средств разработчика F12 в Windows 10</span><span class="sxs-lookup"><span data-stu-id="1e268-102">Debug add-ins using F12 developer tools on Windows 10</span></span>

<span data-ttu-id="1e268-p101">Средства разработчика F12 в Windows 10 помогают отлаживать, тестировать и ускорять веб-страницы. Их также можно использовать для разработки и отладки надстроек Office, если не используется интегрированная среда разработки, например Visual Studio, или если необходимо изучить проблему, запустив надстройку вне интегрированной среды разработки. Средства разработчика F12 можно запускать после запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="1e268-p101">The F12 developer tools included in Windows 10 help you debug, test, and speed up your webpages. You can also use them to develop and debug Office Add-ins, if you are not using an IDE like Visual Studio, or if you need to investigate a problem while running your add-in outside the IDE. You can start the F12 developer tools after your add-in is running.</span></span>

<span data-ttu-id="1e268-p102">В этой статье показано, как использовать отладчик из средств разработчика F12 в Windows 10 для тестирования надстройки Office. Вы можете тестировать надстройки из AppSource или других источников. Средства F12 отображаются в отдельном окне и не используют Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="1e268-p102">This article shows how you how to use the Debugger tool from the F12 developer tools in Windows 10 to test your Office Add-in. You can test add-ins from AppSource or add-ins that you have added from other locations. The F12 tools display in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="1e268-p103">Отладчик входит в состав средств разработчика F12 в Internet Explorer и Windows 10, но не включен в предыдущие версии Windows.</span><span class="sxs-lookup"><span data-stu-id="1e268-p103">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="1e268-111">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="1e268-111">Prerequisites</span></span>

<span data-ttu-id="1e268-112">Вам понадобится следующее программное обеспечение:</span><span class="sxs-lookup"><span data-stu-id="1e268-112">You need the following software:</span></span>

- <span data-ttu-id="1e268-113">средства разработчика F12, которые входят в состав Windows 10;</span><span class="sxs-lookup"><span data-stu-id="1e268-113">The F12 developer tools, which are included in Windows 10.</span></span> 
    
- <span data-ttu-id="1e268-114">клиентское приложение Office, в котором размещается надстройка;</span><span class="sxs-lookup"><span data-stu-id="1e268-114">The Office client application that hosts your add-in.</span></span> 
    
- <span data-ttu-id="1e268-115">надстройка.</span><span class="sxs-lookup"><span data-stu-id="1e268-115">Your add-in.</span></span> 

## <a name="using-the-debugger"></a><span data-ttu-id="1e268-116">Использование отладчика</span><span class="sxs-lookup"><span data-stu-id="1e268-116">Using the Debugger</span></span>

<span data-ttu-id="1e268-117">В этом примере используются Word и бесплатная надстройка из AppSource.</span><span class="sxs-lookup"><span data-stu-id="1e268-117">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="1e268-118">Откройте Word и выберите пустой документ.</span><span class="sxs-lookup"><span data-stu-id="1e268-118">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="1e268-p104">На вкладке **Вставка** в группе "Надстройки" выберите **Магазин**, затем выберите надстройку QR4Office. (Вы можете загрузить любую надстройку из Магазина или каталога надстроек.)</span><span class="sxs-lookup"><span data-stu-id="1e268-p104">On the **Insert** tab, in the Add-ins group, choose **Store** and select the QR4Office add-in. (You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="1e268-121">Запустите средства разработчика F12, которые соответствуют вашей версии Office.</span><span class="sxs-lookup"><span data-stu-id="1e268-121">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="1e268-122">Путь к файлу для 32-разрядной версии Office — C:\Windows\System32\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="1e268-122">For the 32-bit version of Office, use C:\Windows\System32\F12\F12Chooser.exe</span></span>
    
   - <span data-ttu-id="1e268-123">Путь к файлу для 64-разрядной версии Office — C:\Windows\SysWOW64\F12\IEChooser.exe.</span><span class="sxs-lookup"><span data-stu-id="1e268-123">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\F12Chooser.exe</span></span>
    
   <span data-ttu-id="1e268-124">Когда вы запустите IEChooser, в отдельном окне "Выбрать цель для отладки" отобразятся приложения, которые, возможно, нужно отладить.</span><span class="sxs-lookup"><span data-stu-id="1e268-124">When you launch F12Chooser, a separate window named "Choose target to debug" displays the possible applications to debug.</span></span> <span data-ttu-id="1e268-125">Выберите необходимое приложение.</span><span class="sxs-lookup"><span data-stu-id="1e268-125">Select the application that you are interested in.</span></span> <span data-ttu-id="1e268-126">Если вы создаете собственную надстройку, выберите веб-сайт, на котором она развернута. Это может быть URL-адрес localhost.</span><span class="sxs-lookup"><span data-stu-id="1e268-126">If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="1e268-127">Например, выберите **home.html**.</span><span class="sxs-lookup"><span data-stu-id="1e268-127">For example, select **home.html**.</span></span> 
    
   ![Экран IEChooser с указанием на выноску надстройки](../images/choose-target-to-debug.png)

4. <span data-ttu-id="1e268-129">В окне F12 выберите файл, который требуется отладить.</span><span class="sxs-lookup"><span data-stu-id="1e268-129">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="1e268-p106">Чтобы выбрать файл, нажмите значок папки над областью **сценариев** (слева). В раскрывающемся списке появятся доступные файлы. Выберите файл home.js.</span><span class="sxs-lookup"><span data-stu-id="1e268-p106">To select the file, choose the folder icon above the  **script** (left) pane. The dropdown list shows the available files. Select home.js.</span></span>
    
5. <span data-ttu-id="1e268-133">Задайте точку останова.</span><span class="sxs-lookup"><span data-stu-id="1e268-133">Set the breakpoint.</span></span>
    
   <span data-ttu-id="1e268-p107">Чтобы задать точку останова в home.js, выберите строку 144 в функции _textChanged_. Слева от строки появится красная точка, и соответствующая строка отобразится в **области стека вызовов и точек** (справа внизу). Другие способы задания точки останова см. в статье [Проверка выполнения кода JavaScript с помощью отладчика](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span><span class="sxs-lookup"><span data-stu-id="1e268-p107">To set the breakpoint in home.js, choose line 144, which is in the  _textChanged_ function. You will see a red dot to the left of the line and a corresponding line in the **Callstack and Breakpoints** (bottom right) pane. For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span></span> 
    
   ![Отладчик с точкой останова в файле home.js](../images/debugger-home-js-02.png)

6. <span data-ttu-id="1e268-138">Запустите надстройку, чтобы активировать точку останова.</span><span class="sxs-lookup"><span data-stu-id="1e268-138">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="1e268-p108">Выберите текстовое поле URL-адреса в верхней части области QR4Office, чтобы изменить текст. В области **стека вызовов и точек останова** отладчика будет показано, что точка останова активирована и отображает различные сведения. Чтобы увидеть результаты, может потребоваться обновить страницу средства F12.</span><span class="sxs-lookup"><span data-stu-id="1e268-p108">Choose the URL textbox in the upper part of the QR4Office pane to change the text. In the Debugger, in the **Callstack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information. You might need to refresh the F12 tool to see the results.</span></span>
    
   ![Отладчик с результатами из сработавшей точки останова](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="1e268-143">См. также</span><span class="sxs-lookup"><span data-stu-id="1e268-143">See also</span></span>

- <span data-ttu-id="1e268-144">[Проверка выполнения кода JavaScript с помощью отладчика](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="1e268-144">[Inspect running JavaScript with the Debugger](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span></span>
- <span data-ttu-id="1e268-145">[Использование средств разработчика F12](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="1e268-145">[Using the F12 developer tools](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span></span>
    
