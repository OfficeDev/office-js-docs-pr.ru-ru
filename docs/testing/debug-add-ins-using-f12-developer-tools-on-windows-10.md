---
title: Отладка надстроек с помощью средств разработчика в Windows 10
description: ''
ms.date: 07/01/2019
localization_priority: Priority
ms.openlocfilehash: a2090eca41f59f0e7fab1a172aff96cbbca28ed7
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454883"
---
# <a name="debug-add-ins-using-developer-tools-on-windows-10"></a><span data-ttu-id="de05f-102">Отладка надстроек с помощью средств разработчика в Windows 10</span><span class="sxs-lookup"><span data-stu-id="de05f-102">Debug add-ins using F12 developer tools on Windows 10</span></span>

<span data-ttu-id="de05f-103">Для помощи в отладке надстроек в Windows 10 доступны инструменты для разработчиков, не входящие в интегрированные среды разработки.</span><span class="sxs-lookup"><span data-stu-id="de05f-103">There are developer tools outside of IDEs available to help you debug your add-ins on Windows 10.</span></span> <span data-ttu-id="de05f-104">Эти инструменты полезны, если нужно изучить проблемы при запуске настройки вне интегрированной среды разработки.</span><span class="sxs-lookup"><span data-stu-id="de05f-104">These are useful when you need to investigate a problem while running your add-in outside the IDE.</span></span>

<span data-ttu-id="de05f-105">Используемый инструмент зависит от того, где работает настройка: в Microsoft Edge или в Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="de05f-105">The tool that you use depends on whether the add-in is running in Edge or Internet Explorer.</span></span> <span data-ttu-id="de05f-106">Это, в свою очередь, зависит от версии Windows 10 и от версии Office, установленных на компьютере.</span><span class="sxs-lookup"><span data-stu-id="de05f-106">This is determined by the version of Windows 10 and the version of Office that are installed on the computer.</span></span> <span data-ttu-id="de05f-107">Сведения об определении браузера, используемого на компьютере разработки, см. в статье [Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="de05f-107">To determine which browser is being used on your development computer, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span> 


> [!NOTE]
> <span data-ttu-id="de05f-108">Инструкции, представленные в этой статье, нельзя применять для отладки надстройки Outlook, использующей функции Execute.</span><span class="sxs-lookup"><span data-stu-id="de05f-108">The instructions in this article cannot be used to debug an Outlook add-in that uses Execute Functions.</span></span> <span data-ttu-id="de05f-109">Для отладки надстройки Outlook, использующей функции Execute, рекомендуется прикрепить ее к Visual Studio в режиме сценария или к другому отладчику сценариев.</span><span class="sxs-lookup"><span data-stu-id="de05f-109">To debug an Outlook add-in that uses Execute Functions, we recommend that you attach to Visual Studio in script mode or to some other script debugger.</span></span>

## <a name="when-the-add-in-is-running-in-edge"></a><span data-ttu-id="de05f-110">Если надстройка работает в Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="de05f-110">When the add-in is running in Edge</span></span>

<span data-ttu-id="de05f-111">Если надстройка работает в Microsoft Edge, можно использовать [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab).</span><span class="sxs-lookup"><span data-stu-id="de05f-111">When the add-in is running in Edge, you can use the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab).</span></span> 

1. <span data-ttu-id="de05f-112">Запустите надстройку.</span><span class="sxs-lookup"><span data-stu-id="de05f-112">Run the add-in</span></span> 

2. <span data-ttu-id="de05f-113">Запустите Microsoft Edge DevTools.</span><span class="sxs-lookup"><span data-stu-id="de05f-113">Run the Microsoft Edge DevTools.</span></span>

3. <span data-ttu-id="de05f-114">Перейдите на вкладку **Локальные**. Имя вашей надстройки будет указано в списке.</span><span class="sxs-lookup"><span data-stu-id="de05f-114">In the tools, open the **Local** tab. Your add-in will be listed by its name.</span></span>

4. <span data-ttu-id="de05f-115">Щелкните имя надстройки, чтобы открыть ее.</span><span class="sxs-lookup"><span data-stu-id="de05f-115">Click the add-in name to open it in the tools.</span></span>

5. <span data-ttu-id="de05f-116">Перейдите на вкладку **Отладчик**.</span><span class="sxs-lookup"><span data-stu-id="de05f-116">Open the **Permissions** tab.</span></span> 

6. <span data-ttu-id="de05f-117">Выберите значок папки над областью **сценариев** (слева).</span><span class="sxs-lookup"><span data-stu-id="de05f-117">To select the file, choose the folder icon above the  **script** (left) pane.</span></span> <span data-ttu-id="de05f-118">В раскрывающемся списке доступных файлов выберите файл JavaScript, который нужно отладить.</span><span class="sxs-lookup"><span data-stu-id="de05f-118">From the list of available files shown in the dropdown list, select the JavaScript file that you want to debug.</span></span>

7. <span data-ttu-id="de05f-119">Чтобы задать точку останова, выберите строку.</span><span class="sxs-lookup"><span data-stu-id="de05f-119">To set a breakpoint, select the line.</span></span> <span data-ttu-id="de05f-120">Появится красная точка слева от строки и соответствующая строка в области **стека вызовов** (в правом нижнем углу).</span><span class="sxs-lookup"><span data-stu-id="de05f-120">You will see a red dot to the left of the line and a corresponding line in the **Call stack and Breakpoints** (bottom right) pane.</span></span>

8. <span data-ttu-id="de05f-121">Выполните функции в надстройке, необходимые для срабатывания точки останова.</span><span class="sxs-lookup"><span data-stu-id="de05f-121">Execute functions in the add-in as needed to trigger the breakpoint.</span></span>

## <a name="when-the-add-in-is-running-in-internet-explorer"></a><span data-ttu-id="de05f-122">Если надстройка работает в Internet Explorer</span><span class="sxs-lookup"><span data-stu-id="de05f-122">When the add-in is running in Internet Explorer</span></span>

<span data-ttu-id="de05f-123">Если надстройка работает в браузере Internet Explorer, для ее тестирования можно использовать отладчик в составе средств разработчика F12 в Windows 10.</span><span class="sxs-lookup"><span data-stu-id="de05f-123">When the add-in is running in Internet Explorer, you can use the debugger from the F12 developer tools in Windows 10 to test your add-in.</span></span> <span data-ttu-id="de05f-124">Средства разработчика F12 можно запустить после запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="de05f-124">You can start the F12 developer tools after your add-in is running.</span></span> <span data-ttu-id="de05f-125">Средства F12 отображаются в отдельном окне и не используют Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="de05f-125">The F12 tools display in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="de05f-p107">Отладчик входит в состав средств разработчика F12 в Internet Explorer и Windows 10, но не включен в предыдущие версии Windows.</span><span class="sxs-lookup"><span data-stu-id="de05f-p107">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

<span data-ttu-id="de05f-128">В этом примере используются Word и бесплатная надстройка из AppSource.</span><span class="sxs-lookup"><span data-stu-id="de05f-128">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="de05f-129">Откройте Word и выберите пустой документ. </span><span class="sxs-lookup"><span data-stu-id="de05f-129">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="de05f-130">На вкладке **Вставка**, в группе "Надстройки" нажмите **Магазин** и выберите надстройку **QR4Office**.</span><span class="sxs-lookup"><span data-stu-id="de05f-130">On the **Insert** tab, in the Add-ins group, choose **Store** and select the **QR4Office** Add-in.</span></span> <span data-ttu-id="de05f-131">(Вы можете загрузить любую надстройку из Магазина Microsoft Store или каталога надстроек).</span><span class="sxs-lookup"><span data-stu-id="de05f-131">(You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="de05f-132">Запустите средства разработчика F12, которые соответствуют вашей версии Office.</span><span class="sxs-lookup"><span data-stu-id="de05f-132">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="de05f-133">Путь к файлу для 32-разрядной версии Office — C:\Windows\System32\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="de05f-133">For the 32-bit version of Office, use C:\Windows\System32\F12\IEChooser.exe</span></span>
    
   - <span data-ttu-id="de05f-134">Путь к файлу для 64-разрядной версии Office — C:\Windows\SysWOW64\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="de05f-134">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\IEChooser.exe</span></span>
    
   <span data-ttu-id="de05f-135">Когда вы запустите IEChooser, в отдельном окне "Выбрать цель для отладки" отобразятся приложения, которые, возможно, нужно отладить.</span><span class="sxs-lookup"><span data-stu-id="de05f-135">When you launch IEChooser, a separate window named "Choose target to debug" displays the possible applications to debug.</span></span> <span data-ttu-id="de05f-136">Выберите необходимое приложение.</span><span class="sxs-lookup"><span data-stu-id="de05f-136">Select the application that you are interested in.</span></span> <span data-ttu-id="de05f-137">Если вы создаете собственную надстройку, выберите веб-сайт, на котором она развернута. Это может быть URL-адрес localhost.</span><span class="sxs-lookup"><span data-stu-id="de05f-137">If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="de05f-138">Например, выберите **home.html**.</span><span class="sxs-lookup"><span data-stu-id="de05f-138">For example, select **home.html**.</span></span> 
    
   ![Экран IEChooser с выделенной надстройкой](../images/choose-target-to-debug.png)

4. <span data-ttu-id="de05f-140">В окне F12 выберите файл, который требуется отладить.</span><span class="sxs-lookup"><span data-stu-id="de05f-140">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="de05f-141">Чтобы выбрать файл в окне F12, нажмите значок папки над областью **сценариев** (слева).</span><span class="sxs-lookup"><span data-stu-id="de05f-141">To select the file in the F12 window, choose the folder icon above the **script** (left) pane.</span></span> <span data-ttu-id="de05f-142">В списке доступных файлов, представленных в раскрывающемся списке, выберите **Home.js**.</span><span class="sxs-lookup"><span data-stu-id="de05f-142">From the list of available files shown in the dropdown list, select **Home.js**.</span></span>
    
5. <span data-ttu-id="de05f-143">Задайте точку останова.</span><span class="sxs-lookup"><span data-stu-id="de05f-143">Set the breakpoint.</span></span>
    
   <span data-ttu-id="de05f-144">Чтобы задать точку останова в файле **Home.js**, выберите строку 144 (код функции `textChanged`).</span><span class="sxs-lookup"><span data-stu-id="de05f-144">To set the breakpoint in **Home.js**, choose line 144, which is in the  `textChanged` function.</span></span> <span data-ttu-id="de05f-145">Появятся красная точка слева от строки и соответствующая строка в области стека вызовов и точек останова (справа внизу).</span><span class="sxs-lookup"><span data-stu-id="de05f-145">You will see a red dot to the left of the line and a corresponding line in the **Call stack and Breakpoints** (bottom right) pane.</span></span> <span data-ttu-id="de05f-146">Другие способы задания точки останова см. в статье [Проверка выполнения кода JavaScript при помощи отладчика](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span><span class="sxs-lookup"><span data-stu-id="de05f-146">For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span></span> 
    
   ![Отладчик с точкой останова в файле home.js](../images/debugger-home-js-02.png)

6. <span data-ttu-id="de05f-148">Запустите надстройку, чтобы активировать точку останова.</span><span class="sxs-lookup"><span data-stu-id="de05f-148">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="de05f-149">В Word выберите текстовое поле URL-адреса в верхней части области **QR4Office** и попробуйте ввести какой-либо текст.</span><span class="sxs-lookup"><span data-stu-id="de05f-149">In Word, choose the URL textbox in the upper part of the **QR4Office** pane and attempt to enter some text.</span></span> <span data-ttu-id="de05f-150">В области **стека вызовов и точек останова** в отладчике вы увидите, что точка останова активирована и показывает различные сведения.</span><span class="sxs-lookup"><span data-stu-id="de05f-150">In the Debugger, in the **Call stack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information.</span></span> <span data-ttu-id="de05f-151">Чтобы увидеть результаты, может потребоваться обновить отладчик.</span><span class="sxs-lookup"><span data-stu-id="de05f-151">You might need to refresh the Debugger to see the results.</span></span>
    
   ![Отладчик с результатами из сработавшей точки останова](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="de05f-153">См. также</span><span class="sxs-lookup"><span data-stu-id="de05f-153">See also</span></span>

- <span data-ttu-id="de05f-154">[Проверка выполнения кода JavaScript с помощью отладчика](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="de05f-154">[Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span></span>
- <span data-ttu-id="de05f-155">[Использование средств разработчика F12](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="de05f-155">[Using the F12 developer tools](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span></span>
