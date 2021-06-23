---
title: Отладка надстроек с помощью средств разработчика в Windows 10
description: Отладка надстроек с помощью средств разработчика Microsoft Edge в Windows 10
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 41e7f2c8efb6406948c30522b56424ed7f9aa400
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076534"
---
# <a name="debug-add-ins-using-developer-tools-on-windows-10"></a><span data-ttu-id="018da-103">Отладка надстроек с помощью средств разработчика в Windows 10</span><span class="sxs-lookup"><span data-stu-id="018da-103">Debug add-ins using developer tools on Windows 10</span></span>

<span data-ttu-id="018da-104">Для помощи в отладке надстроек в Windows 10 доступны инструменты для разработчиков, не входящие в интегрированные среды разработки.</span><span class="sxs-lookup"><span data-stu-id="018da-104">There are developer tools outside of IDEs available to help you debug your add-ins on Windows 10.</span></span> <span data-ttu-id="018da-105">Эти инструменты полезны, если нужно изучить проблемы при запуске надстройки вне интегрированной среды разработки.</span><span class="sxs-lookup"><span data-stu-id="018da-105">These are useful when you need to investigate a problem while running your add-in outside the IDE.</span></span>

<span data-ttu-id="018da-106">Используемый инструмент зависит от того, где работает надстройка: в Microsoft Edge или в Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="018da-106">The tool that you use depends on whether the add-in is running in Microsoft Edge or Internet Explorer.</span></span> <span data-ttu-id="018da-107">Это, в свою очередь, зависит от версий Windows 10 и Office, установленных на компьютере.</span><span class="sxs-lookup"><span data-stu-id="018da-107">This is determined by the version of Windows 10 and the version of Office that are installed on the computer.</span></span> <span data-ttu-id="018da-108">Сведения об определении браузера, используемого на компьютере разработки, см. в статье [Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="018da-108">To determine which browser is being used on your development computer, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span>

> [!NOTE]
> <span data-ttu-id="018da-109">Инструкции, представленные в этой статье, нельзя применять для отладки надстройки Outlook, использующей функции Execute.</span><span class="sxs-lookup"><span data-stu-id="018da-109">The instructions in this article cannot be used to debug an Outlook add-in that uses Execute Functions.</span></span> <span data-ttu-id="018da-110">Для отладки надстройки Outlook, использующей функции Execute, рекомендуется прикрепить ее к Visual Studio в режиме сценария или к другому отладчику сценариев.</span><span class="sxs-lookup"><span data-stu-id="018da-110">To debug an Outlook add-in that uses Execute Functions, we recommend that you attach to Visual Studio in script mode or to some other script debugger.</span></span>

## <a name="when-the-add-in-is-running-in-microsoft-edge"></a><span data-ttu-id="018da-111">Если надстройка работает в Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="018da-111">When the add-in is running in Microsoft Edge</span></span>

[!include[Enable debugging on Microsoft Edge DevTools](../includes/enable-debugging-on-edge-devtools.md)]

### <a name="debug-using-microsoft-edge-devtools"></a><span data-ttu-id="018da-112">Отладка с помощью Microsoft Edge DevTools</span><span class="sxs-lookup"><span data-stu-id="018da-112">Debug using Microsoft Edge DevTools</span></span>

<span data-ttu-id="018da-113">Если надстройка работает в Microsoft Edge, можно использовать [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab).</span><span class="sxs-lookup"><span data-stu-id="018da-113">When the add-in is running in Microsoft Edge, you can use the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab).</span></span>

1. <span data-ttu-id="018da-114">Запустите надстройку.</span><span class="sxs-lookup"><span data-stu-id="018da-114">Run the add-in.</span></span>

2. <span data-ttu-id="018da-115">Запустите Microsoft Edge DevTools.</span><span class="sxs-lookup"><span data-stu-id="018da-115">Run the Microsoft Edge DevTools.</span></span>

3. <span data-ttu-id="018da-116">Перейдите на вкладку **Локальные**. Имя вашей надстройки будет указано в списке.</span><span class="sxs-lookup"><span data-stu-id="018da-116">In the tools, open the **Local** tab. Your add-in will be listed by its name.</span></span>

4. <span data-ttu-id="018da-117">Щелкните имя надстройки, чтобы открыть ее.</span><span class="sxs-lookup"><span data-stu-id="018da-117">Click the add-in name to open it in the tools.</span></span>

5. <span data-ttu-id="018da-118">Перейдите на вкладку **Отладчик**.</span><span class="sxs-lookup"><span data-stu-id="018da-118">Open the **Debugger** tab.</span></span> 

6. <span data-ttu-id="018da-119">Выберите значок папки над областью **сценариев** (слева).</span><span class="sxs-lookup"><span data-stu-id="018da-119">Choose the folder icon above the **script** (left) pane.</span></span> <span data-ttu-id="018da-120">В раскрывающемся списке доступных файлов выберите файл JavaScript, который нужно отладить.</span><span class="sxs-lookup"><span data-stu-id="018da-120">From the list of available files shown in the dropdown list, select the JavaScript file that you want to debug.</span></span>

7. <span data-ttu-id="018da-121">Чтобы задать точку останова, выберите строку.</span><span class="sxs-lookup"><span data-stu-id="018da-121">To set a breakpoint, select the line.</span></span> <span data-ttu-id="018da-122">Появится красная точка слева от строки и соответствующая строка в области **стека вызовов** (в правом нижнем углу).</span><span class="sxs-lookup"><span data-stu-id="018da-122">You will see a red dot to the left of the line and a corresponding line in the **Call stack** (bottom right) pane.</span></span>

8. <span data-ttu-id="018da-123">Выполните функции в надстройке, необходимые для срабатывания точки останова.</span><span class="sxs-lookup"><span data-stu-id="018da-123">Execute functions in the add-in as needed to trigger the breakpoint.</span></span>

## <a name="when-the-add-in-is-running-in-internet-explorer"></a><span data-ttu-id="018da-124">Если надстройка работает в Internet Explorer</span><span class="sxs-lookup"><span data-stu-id="018da-124">When the add-in is running in Internet Explorer</span></span>

<span data-ttu-id="018da-125">Если надстройка работает в браузере Internet Explorer, для ее тестирования можно использовать отладчик в составе средств разработчика F12 в Windows 10.</span><span class="sxs-lookup"><span data-stu-id="018da-125">When the add-in is running in Internet Explorer, you can use the debugger from the F12 developer tools in Windows 10 to test your add-in.</span></span> <span data-ttu-id="018da-126">Средства разработчика F12 можно запустить после запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="018da-126">You can start the F12 developer tools after the add-in is running.</span></span> <span data-ttu-id="018da-127">Средства F12 отображаются в отдельном окне и не используют Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="018da-127">The F12 tools are displayed in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="018da-p107">Отладчик входит в состав средств разработчика F12 в Internet Explorer и Windows 10, но не включен в предыдущие версии Windows.</span><span class="sxs-lookup"><span data-stu-id="018da-p107">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

<span data-ttu-id="018da-130">В этом примере используются Word и бесплатная надстройка из AppSource.</span><span class="sxs-lookup"><span data-stu-id="018da-130">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="018da-131">Откройте Word и выберите пустой документ. </span><span class="sxs-lookup"><span data-stu-id="018da-131">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="018da-132">На вкладке **Вставка**, в группе "Надстройки" нажмите **Магазин** и выберите надстройку **QR4Office**.</span><span class="sxs-lookup"><span data-stu-id="018da-132">On the **Insert** tab, in the Add-ins group, choose **Store** and select the **QR4Office** Add-in.</span></span> <span data-ttu-id="018da-133">(Вы можете загрузить любую надстройку из Магазина Microsoft Store или каталога надстроек).</span><span class="sxs-lookup"><span data-stu-id="018da-133">(You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="018da-134">Запустите средства разработчика F12, которые соответствуют вашей версии Office.</span><span class="sxs-lookup"><span data-stu-id="018da-134">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="018da-135">Путь к файлу для 32-разрядной версии Office — C:\Windows\System32\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="018da-135">For the 32-bit version of Office, use C:\Windows\System32\F12\IEChooser.exe</span></span>
    
   - <span data-ttu-id="018da-136">Путь к файлу для 64-разрядной версии Office — C:\Windows\SysWOW64\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="018da-136">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\IEChooser.exe</span></span>
    
   <span data-ttu-id="018da-137">Когда вы запустите IEChooser, в отдельном окне "Выбрать цель для отладки" отобразятся приложения, которые, возможно, нужно отладить.</span><span class="sxs-lookup"><span data-stu-id="018da-137">When you launch IEChooser, a separate window named "Choose target to debug" displays the possible applications to debug.</span></span> <span data-ttu-id="018da-138">Выберите необходимое приложение.</span><span class="sxs-lookup"><span data-stu-id="018da-138">Select the application that you are interested in.</span></span> <span data-ttu-id="018da-139">Если вы создаете собственную надстройку, выберите веб-сайт, на котором она развернута. Это может быть URL-адрес localhost.</span><span class="sxs-lookup"><span data-stu-id="018da-139">If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="018da-140">Например, выберите **home.html**.</span><span class="sxs-lookup"><span data-stu-id="018da-140">For example, select **home.html**.</span></span> 
    
   ![Экран IEChooser, указывающий на надстройку пузырьков.](../images/choose-target-to-debug.png)

4. <span data-ttu-id="018da-142">В окне F12 выберите файл, который требуется отладить.</span><span class="sxs-lookup"><span data-stu-id="018da-142">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="018da-143">Чтобы выбрать файл в окне F12, нажмите значок папки над областью **сценариев** (слева).</span><span class="sxs-lookup"><span data-stu-id="018da-143">To select the file in the F12 window, choose the folder icon above the **script** (left) pane.</span></span> <span data-ttu-id="018da-144">В списке доступных файлов, представленных в раскрывающемся списке, выберите **Home.js**.</span><span class="sxs-lookup"><span data-stu-id="018da-144">From the list of available files shown in the dropdown list, select **Home.js**.</span></span>
    
5. <span data-ttu-id="018da-145">Задайте точку останова.</span><span class="sxs-lookup"><span data-stu-id="018da-145">Set the breakpoint.</span></span>
    
   <span data-ttu-id="018da-146">Чтобы задать точку останова в файле **Home.js**, выберите строку 144 (код функции `textChanged`).</span><span class="sxs-lookup"><span data-stu-id="018da-146">To set the breakpoint in **Home.js**, choose line 144, which is in the  `textChanged` function.</span></span> <span data-ttu-id="018da-147">Появятся красная точка слева от строки и соответствующая строка в области стека вызовов и точек останова (справа внизу).</span><span class="sxs-lookup"><span data-stu-id="018da-147">You will see a red dot to the left of the line and a corresponding line in the **Call stack and Breakpoints** (bottom right) pane.</span></span> <span data-ttu-id="018da-148">Другие способы задания точки останова см. в статье [Проверка выполнения кода JavaScript при помощи отладчика](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span><span class="sxs-lookup"><span data-stu-id="018da-148">For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span></span> 
    
   ![Отладка с брейк-пойнтом в home.js файле.](../images/debugger-home-js-02.png)

6. <span data-ttu-id="018da-150">Запустите надстройку, чтобы активировать точку останова.</span><span class="sxs-lookup"><span data-stu-id="018da-150">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="018da-151">В Word выберите текстовое поле URL-адреса в верхней части области **QR4Office** и попробуйте ввести какой-либо текст.</span><span class="sxs-lookup"><span data-stu-id="018da-151">In Word, choose the URL textbox in the upper part of the **QR4Office** pane and attempt to enter some text.</span></span> <span data-ttu-id="018da-152">В области **стека вызовов и точек останова** в отладчике вы увидите, что точка останова активирована и показывает различные сведения.</span><span class="sxs-lookup"><span data-stu-id="018da-152">In the Debugger, in the **Call stack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information.</span></span> <span data-ttu-id="018da-153">Чтобы увидеть результаты, может потребоваться обновить отладчик.</span><span class="sxs-lookup"><span data-stu-id="018da-153">You might need to refresh the Debugger to see the results.</span></span>
    
   ![Отладка результатов с срабатывуемой точкой разрыва.](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="018da-155">См. также</span><span class="sxs-lookup"><span data-stu-id="018da-155">See also</span></span>

- <span data-ttu-id="018da-156">[Проверка выполнения кода JavaScript с помощью отладчика](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="018da-156">[Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span></span>
- <span data-ttu-id="018da-157">[Использование средств разработчика F12](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="018da-157">[Using the F12 developer tools](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span></span>
