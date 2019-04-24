---
title: Отладка надстроек с помощью средств разработчика F12 в Windows 10
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 750411bea187a0ade9b3723e3198d82f7c482c9f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450155"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a><span data-ttu-id="b81ad-102">Отладка надстроек с помощью средств разработчика F12 в Windows 10</span><span class="sxs-lookup"><span data-stu-id="b81ad-102">Debug add-ins using F12 developer tools on Windows 10</span></span>

<span data-ttu-id="b81ad-103">Средства разработчика F12 в Windows 10 помогают отлаживать, тестировать и ускорять веб-страницы.</span><span class="sxs-lookup"><span data-stu-id="b81ad-103">The F12 developer tools included in Windows 10 help you debug, test, and speed up your webpages.</span></span> <span data-ttu-id="b81ad-104">Их также можно использовать для разработки и отладки надстроек Office, если не используется интегрированная среда разработки, например Visual Studio, или если необходимо изучить проблему, запустив надстройку вне интегрированной среды разработки.</span><span class="sxs-lookup"><span data-stu-id="b81ad-104">You can also use them to develop and debug Office Add-ins, if you are not using an IDE like Visual Studio, or if you need to investigate a problem while running your add-in outside the IDE.</span></span> <span data-ttu-id="b81ad-105">В этой статье описано, как использовать отладчик из средств разработчика F12 в Windows 10 для тестирования надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="b81ad-105">This article describes how to use the Debugger tool from the F12 developer tools in Windows 10 to test your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="b81ad-106">Инструкции, представленные в этой статье, нельзя применять для отладки надстройки Outlook, использующей функции Execute.</span><span class="sxs-lookup"><span data-stu-id="b81ad-106">The instructions in this article cannot be used to debug an Outlook add-in that uses Execute Functions.</span></span> <span data-ttu-id="b81ad-107">Для отладки надстройки Outlook, использующей функции Execute, рекомендуется прикрепить ее к Visual Studio в режиме сценария или к другому отладчику сценариев.</span><span class="sxs-lookup"><span data-stu-id="b81ad-107">To debug an Outlook add-in that uses Execute Functions, we recommend that you attach to Visual Studio in script mode or to some other script debugger.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="b81ad-108">Предварительные условия</span><span class="sxs-lookup"><span data-stu-id="b81ad-108">Prerequisites</span></span>

<span data-ttu-id="b81ad-109">Вам понадобится следующее программное обеспечение:</span><span class="sxs-lookup"><span data-stu-id="b81ad-109">You need the following software:</span></span>

- <span data-ttu-id="b81ad-110">средства разработчика F12, которые входят в состав Windows 10; </span><span class="sxs-lookup"><span data-stu-id="b81ad-110">The F12 developer tools, which are included in Windows 10.</span></span> 
    
- <span data-ttu-id="b81ad-111">клиентское приложение Office, в котором размещается надстройка; </span><span class="sxs-lookup"><span data-stu-id="b81ad-111">The Office client application that hosts your add-in.</span></span> 
    
- <span data-ttu-id="b81ad-112">надстройка. </span><span class="sxs-lookup"><span data-stu-id="b81ad-112">Your add-in.</span></span> 

## <a name="using-the-debugger"></a><span data-ttu-id="b81ad-113">Использование отладчика</span><span class="sxs-lookup"><span data-stu-id="b81ad-113">Using the Debugger</span></span>

<span data-ttu-id="b81ad-114">В этой статье показано, как использовать отладчик из средств разработчика F12 в Windows 10 для тестирования надстройки Office. Вы можете тестировать надстройки из AppSource или других источников. Средства F12 отображаются в отдельном окне и не используют Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="b81ad-114">You can use the Debugger from the F12 developer tools in Windows 10 to test add-ins from AppSource or add-ins that you have added from other locations.</span></span> <span data-ttu-id="b81ad-115">Средства разработчика F12 можно запускать после запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="b81ad-115">You can start the F12 developer tools after your add-in is running.</span></span> <span data-ttu-id="b81ad-116">Средства F12 отображаются в отдельном окне и не используют Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="b81ad-116">The F12 tools display in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="b81ad-p104">Отладчик входит в состав средств разработчика F12 в Internet Explorer и Windows 10, но не включен в предыдущие версии Windows.</span><span class="sxs-lookup"><span data-stu-id="b81ad-p104">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

<span data-ttu-id="b81ad-119">В этом примере используются Word и бесплатная надстройка из AppSource.</span><span class="sxs-lookup"><span data-stu-id="b81ad-119">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="b81ad-120">Откройте Word и выберите пустой документ. </span><span class="sxs-lookup"><span data-stu-id="b81ad-120">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="b81ad-121">На вкладке **Вставка**, в группе "Надстройки" нажмите **Магазин** и выберите надстройку **QR4Office**.</span><span class="sxs-lookup"><span data-stu-id="b81ad-121">On the **Insert** tab, in the Add-ins group, choose **Store** and select the **QR4Office** Add-in.</span></span> <span data-ttu-id="b81ad-122">(Вы можете загрузить любую надстройку из Магазина Microsoft Store или каталога надстроек).</span><span class="sxs-lookup"><span data-stu-id="b81ad-122">(You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="b81ad-123">Запустите средства разработчика F12, которые соответствуют вашей версии Office.</span><span class="sxs-lookup"><span data-stu-id="b81ad-123">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="b81ad-124">Путь к файлу для 32-разрядной версии Office — C:\Windows\System32\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="b81ad-124">For the 32-bit version of Office, use C:\Windows\System32\F12\IEChooser.exe</span></span>
    
   - <span data-ttu-id="b81ad-125">Путь к файлу для 64-разрядной версии Office — C:\Windows\SysWOW64\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="b81ad-125">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\IEChooser.exe</span></span>
    
   <span data-ttu-id="b81ad-126">Когда вы запустите IEChooser, в отдельном окне "Выбрать цель для отладки" отобразятся приложения, которые, возможно, нужно отладить.</span><span class="sxs-lookup"><span data-stu-id="b81ad-126">When you launch IEChooser, a separate window named "Choose target to debug" displays the possible applications to debug.</span></span> <span data-ttu-id="b81ad-127">Выберите необходимое приложение.</span><span class="sxs-lookup"><span data-stu-id="b81ad-127">Select the application that you are interested in.</span></span> <span data-ttu-id="b81ad-128">Если вы создаете собственную надстройку, выберите веб-сайт, на котором она развернута. Это может быть URL-адрес localhost.</span><span class="sxs-lookup"><span data-stu-id="b81ad-128">If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="b81ad-129">Например, выберите **home.html**.</span><span class="sxs-lookup"><span data-stu-id="b81ad-129">For example, select **home.html**.</span></span> 
    
   ![Экран IEChooser с выделенной надстройкой](../images/choose-target-to-debug.png)

4. <span data-ttu-id="b81ad-131">В окне F12 выберите файл, который требуется отладить.</span><span class="sxs-lookup"><span data-stu-id="b81ad-131">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="b81ad-132">Чтобы выбрать файл в окне F12, нажмите значок папки над областью **сценариев** (слева).</span><span class="sxs-lookup"><span data-stu-id="b81ad-132">To select the file in the F12 window, choose the folder icon above the **script** (left) pane.</span></span> <span data-ttu-id="b81ad-133">В списке доступных файлов, представленных в раскрывающемся списке, выберите **Home.js**.</span><span class="sxs-lookup"><span data-stu-id="b81ad-133">From the list of available files shown in the dropdown list, select **Home.js**.</span></span>
    
5. <span data-ttu-id="b81ad-134">Задайте точку останова.</span><span class="sxs-lookup"><span data-stu-id="b81ad-134">Set the breakpoint.</span></span>
    
   <span data-ttu-id="b81ad-135">Чтобы задать точку останова в файле **Home.js**, выберите строку 144 (код функции `textChanged`).</span><span class="sxs-lookup"><span data-stu-id="b81ad-135">To set the breakpoint in **Home.js**, choose line 144, which is in the  `textChanged` function.</span></span> <span data-ttu-id="b81ad-136">Появятся красная точка слева от строки и соответствующая строка в области стека вызовов и точек останова (справа внизу).</span><span class="sxs-lookup"><span data-stu-id="b81ad-136">You will see a red dot to the left of the line and a corresponding line in the **Call stack and Breakpoints** (bottom right) pane.</span></span> <span data-ttu-id="b81ad-137">Другие способы задания точки останова см. в статье [Проверка выполнения кода JavaScript при помощи отладчика](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span><span class="sxs-lookup"><span data-stu-id="b81ad-137">For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span></span> 
    
   ![Отладчик с точкой останова в файле home.js](../images/debugger-home-js-02.png)

6. <span data-ttu-id="b81ad-139">Запустите надстройку, чтобы активировать точку останова.</span><span class="sxs-lookup"><span data-stu-id="b81ad-139">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="b81ad-140">В Word выберите текстовое поле URL-адреса в верхней части области **QR4Office** и попробуйте ввести какой-либо текст.</span><span class="sxs-lookup"><span data-stu-id="b81ad-140">In Word, choose the URL textbox in the upper part of the **QR4Office** pane and attempt to enter some text.</span></span> <span data-ttu-id="b81ad-141">В области **стека вызовов и точек останова** в отладчике вы увидите, что точка останова активирована и показывает различные сведения.</span><span class="sxs-lookup"><span data-stu-id="b81ad-141">In the Debugger, in the **Call stack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information.</span></span> <span data-ttu-id="b81ad-142">Чтобы увидеть результаты, может потребоваться обновить отладчик.</span><span class="sxs-lookup"><span data-stu-id="b81ad-142">You might need to refresh the Debugger to see the results.</span></span>
    
   ![Отладчик с результатами из сработавшей точки останова](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="b81ad-144">См. также</span><span class="sxs-lookup"><span data-stu-id="b81ad-144">See also</span></span>

- <span data-ttu-id="b81ad-145">[Проверка выполнения кода JavaScript с помощью отладчика](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="b81ad-145">[Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span></span>
- <span data-ttu-id="b81ad-146">[Использование средств разработчика F12](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="b81ad-146">[Using the F12 developer tools](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span></span>
