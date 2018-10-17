---
title: Отладка надстроек с помощью средств разработчика F12 в Windows 10
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: 3df245fcd651ec227e0a32d53da186ee332beb8f
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579844"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a><span data-ttu-id="bb00b-102">Отладка надстроек с помощью средств разработчика F12 в Windows 10</span><span class="sxs-lookup"><span data-stu-id="bb00b-102">Debug add-ins using F12 developer tools on Windows 10</span></span>

<span data-ttu-id="bb00b-p101">Средства разработчика F12 в Windows 10 помогают отлаживать, тестировать и ускорять веб-страницы. Их также можно использовать для разработки и отладки надстроек Office, если не используется интегрированная среда разработки, например Visual Studio, или если необходимо изучить проблему, запустив надстройку вне интегрированной среды разработки.  В этой статье описывается, как использовать средство отладки из средства разработчика F12 в Windows 10 для тестирования надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="bb00b-p101">The F12 developer tools included in Windows 10 help you debug, test, and speed up your webpages. You can also use them to develop and debug Office Add-ins, if you are not using an IDE like Visual Studio, or if you need to investigate a problem while running your add-in outside the IDE. You can start the F12 developer tools after your add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="bb00b-106">Инструкцию, представленную в этой статье, нельзя использоваться для отладки надстройки Outlook, которая использует выполнение функций.</span><span class="sxs-lookup"><span data-stu-id="bb00b-106">The instructions in this article cannot be used to debug an Outlook add-in that uses Execute Functions.</span></span> <span data-ttu-id="bb00b-107">Для отладки надстройки Outlook, в которой используется выполнение функций, мы рекомендуем прикрепить Visual Studio в режиме сценария или какой-либо другой отладчик сценариев.</span><span class="sxs-lookup"><span data-stu-id="bb00b-107">To debug an Outlook add-in that uses Execute Functions, we recommend that you attach to Visual Studio in script mode or to some other script debugger.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="bb00b-108">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="bb00b-108">Prerequisites</span></span>

<span data-ttu-id="bb00b-109">Вам понадобится следующее программное обеспечение:</span><span class="sxs-lookup"><span data-stu-id="bb00b-109">You need the following software:</span></span>

- <span data-ttu-id="bb00b-110">средства разработчика F12, которые входят в состав Windows 10;</span><span class="sxs-lookup"><span data-stu-id="bb00b-110">The F12 developer tools, which are included in Windows 10.</span></span> 
    
- <span data-ttu-id="bb00b-111">клиентское приложение Office, в котором размещается надстройка;</span><span class="sxs-lookup"><span data-stu-id="bb00b-111">The Office client application that hosts your add-in.</span></span> 
    
- <span data-ttu-id="bb00b-112">ваша надстройка.</span><span class="sxs-lookup"><span data-stu-id="bb00b-112">Your add-in.</span></span> 

## <a name="using-the-debugger"></a><span data-ttu-id="bb00b-113">Использование отладчика</span><span class="sxs-lookup"><span data-stu-id="bb00b-113">Using the Debugger</span></span>

<span data-ttu-id="bb00b-114">Вы можете использовать отладчик от средства разработчика F12 в Windows 10 для тестирования надстроек из AppSource или других источников.</span><span class="sxs-lookup"><span data-stu-id="bb00b-114">This article shows how you how to use the Debugger tool from the F12 developer tools in Windows 10 to test your Office Add-in. You can test add-ins from AppSource or add-ins that you have added from other locations. The F12 tools display in a separate window and do not use Visual Studio.</span></span> <span data-ttu-id="bb00b-115">Работку средства разработчика F12 можно начать после запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="bb00b-115">You can start the F12 developer tools after your add-in is running.</span></span> <span data-ttu-id="bb00b-116">Средства F12 отображаются в отдельном окне и не используют Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="bb00b-116">The F12 tools display in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="bb00b-p104">Отладчик входит в состав средств разработчика F12 в Internet Explorer и Windows 10, но не включен в предыдущие версии Windows.</span><span class="sxs-lookup"><span data-stu-id="bb00b-p104">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

<span data-ttu-id="bb00b-119">В этом примере используются Word и бесплатная надстройка из AppSource.</span><span class="sxs-lookup"><span data-stu-id="bb00b-119">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="bb00b-120">Откройте Word и выберите пустой документ.</span><span class="sxs-lookup"><span data-stu-id="bb00b-120">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="bb00b-121">На вкладке **Вставка** в группе "Надстройки" выберите **Магазин**, затем выберите надстройку **QR4Office**.</span><span class="sxs-lookup"><span data-stu-id="bb00b-121">On the Insert tab, in the Add-ins group, choose Store and select the QR4Office add-in. (You can load any add-in from the Store or your add-in catalog.)</span></span> <span data-ttu-id="bb00b-122">(Вы можете загрузить любую надстройку из Магазина или каталога надстроек.)</span><span class="sxs-lookup"><span data-stu-id="bb00b-122">On the  Insert tab, in the Add-ins group, Store and select the QR4Office add-in. (You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="bb00b-123">Запустите средства разработчика F12, которые соответствуют вашей версии Office.</span><span class="sxs-lookup"><span data-stu-id="bb00b-123">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="bb00b-124">Путь к файлу для 32-разрядной версии Office — C:\Windows\System32\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="bb00b-124">For the 32-bit version of Office, use C:\Windows\System32\F12\F12Chooser.exe</span></span>
    
   - <span data-ttu-id="bb00b-125">Путь к файлу для 64-разрядной версии Office — C:\Windows\SysWOW64\F12\IEChooser.exe.</span><span class="sxs-lookup"><span data-stu-id="bb00b-125">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\F12Chooser.exe</span></span>
    
   <span data-ttu-id="bb00b-126">Когда вы запустите IEChooser, в отдельном окне "Выбрать цель для отладки" отобразятся приложения, которые, возможно, нужно отладить.</span><span class="sxs-lookup"><span data-stu-id="bb00b-126">When you launch F12Chooser, a separate window named "Choose target to debug" displays the possible applications to debug.</span></span> <span data-ttu-id="bb00b-127">Выберите необходимое приложение.</span><span class="sxs-lookup"><span data-stu-id="bb00b-127">Select the application that you are interested in.</span></span> <span data-ttu-id="bb00b-128">Если вы создаете собственную надстройку, выберите веб-сайт, на котором она развернута. Это может быть URL-адрес localhost.</span><span class="sxs-lookup"><span data-stu-id="bb00b-128">If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="bb00b-129">Например, выберите **home.html**.</span><span class="sxs-lookup"><span data-stu-id="bb00b-129">For example, select **home.html**.</span></span> 
    
   ![Экран IEChooser с указанием на выноску надстройки](../images/choose-target-to-debug.png)

4. <span data-ttu-id="bb00b-131">В окне F12 выберите файл, который требуется отладить.</span><span class="sxs-lookup"><span data-stu-id="bb00b-131">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="bb00b-132">Чтобы выбрать файл в онке F12, нажмите значок папки над областью **сценариев** (слева).</span><span class="sxs-lookup"><span data-stu-id="bb00b-132">To select the file, choose the folder icon above the  **script** (left) pane.</span></span> <span data-ttu-id="bb00b-133">В списке доступных файлов, представленных в раскрывающемся списке, выберите **Home.js**.</span><span class="sxs-lookup"><span data-stu-id="bb00b-133">From the list of available files shown in the dropdown list, select **Home.js**.</span></span>
    
5. <span data-ttu-id="bb00b-134">Задайте точку останова.</span><span class="sxs-lookup"><span data-stu-id="bb00b-134">Set the breakpoint.</span></span>
    
   <span data-ttu-id="bb00b-135">Чтобы задать точку останова в файле **Home.js**, выберите строку 144 (код функции `textChanged`).</span><span class="sxs-lookup"><span data-stu-id="bb00b-135">To set the breakpoint in home.js, choose line 144, which is in the  textChanged function.</span></span> <span data-ttu-id="bb00b-136">Появятся красная точка слева от строки и соответствующая строка в области **стека вызовов и точек останова** (справа внизу).</span><span class="sxs-lookup"><span data-stu-id="bb00b-136">You will see a red dot to the left of the line and a corresponding line in the **Callstack and Breakpoints** (bottom right) pane.</span></span> <span data-ttu-id="bb00b-137">Другие способы задания точки останова см. в статье [Проверка выполнения кода JavaScript при помощи отладчика](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span><span class="sxs-lookup"><span data-stu-id="bb00b-137">For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span></span> 
    
   ![Отладчик с точкой останова в файле home.js](../images/debugger-home-js-02.png)

6. <span data-ttu-id="bb00b-139">Запустите надстройку, чтобы активировать точку останова.</span><span class="sxs-lookup"><span data-stu-id="bb00b-139">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="bb00b-140">В Word выберите текстовое поле URL-адреса в верхней части области **QR4Office** и попробуйте ввести какой-либо текст.</span><span class="sxs-lookup"><span data-stu-id="bb00b-140">In Word, choose the URL textbox in the upper part of the **QR4Office** pane and attempt to enter some text.</span></span> <span data-ttu-id="bb00b-141">В области **стека вызовов и точек останова** в отладчике вы увидите, что точка останова активирована и показывает различные сведения.</span><span class="sxs-lookup"><span data-stu-id="bb00b-141">In the Debugger, in the  **Callstack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information.</span></span> <span data-ttu-id="bb00b-142">Чтобы увидеть результаты, может потребоваться обновить средство F12.</span><span class="sxs-lookup"><span data-stu-id="bb00b-142">You might need to refresh the F12 tool to see the results.</span></span>
    
   ![Отладчик с результатами из сработавшей точки останова](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="bb00b-144">См. также</span><span class="sxs-lookup"><span data-stu-id="bb00b-144">See also</span></span>

- <span data-ttu-id="bb00b-145">[Проверка выполнения кода JavaScript с помощью отладчика](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="bb00b-145">[Inspect running JavaScript with the Debugger](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span></span>
- <span data-ttu-id="bb00b-146">[Использование средств разработчика F12](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="bb00b-146">[Using the F12 developer tools](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span></span>
