---
title: Отламывка надстройки Outlook событий (предварительный просмотр)
description: Узнайте, как отлагировать Outlook надстройки, которая реализует активацию на основе событий.
ms.topic: article
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: d7621a7407db3b8e773d1534beb6c881f7b48558
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555285"
---
# <a name="debug-your-event-based-outlook-add-in-preview"></a><span data-ttu-id="9134a-103">Отламывка надстройки Outlook событий (предварительный просмотр)</span><span class="sxs-lookup"><span data-stu-id="9134a-103">Debug your event-based Outlook add-in (preview)</span></span>

<span data-ttu-id="9134a-104">В этой статье содержится руководство по отладки при реализации активации на основе событий [в](autolaunch.md) надстройки.</span><span class="sxs-lookup"><span data-stu-id="9134a-104">This article provides debugging guidance as you implement [event-based activation](autolaunch.md) in your add-in.</span></span> <span data-ttu-id="9134a-105">Функция активации на основе событий в настоящее время находится в предварительном режиме.</span><span class="sxs-lookup"><span data-stu-id="9134a-105">The event-based activation feature is currently in preview.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9134a-106">Эта возможность отладки поддерживается только для предварительного просмотра Outlook в Windows с Microsoft 365 подпиской.</span><span class="sxs-lookup"><span data-stu-id="9134a-106">This debugging capability is only supported for preview in Outlook on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="9134a-107">Дополнительные сведения см. в разделе [Отладка](#preview-debugging-for-the-event-based-activation-feature) предварительного просмотра для раздела функции активации на основе событий в этой статье.</span><span class="sxs-lookup"><span data-stu-id="9134a-107">For more information, see the [Preview debugging for the event-based activation feature](#preview-debugging-for-the-event-based-activation-feature) section in this article.</span></span>

<span data-ttu-id="9134a-108">В этой статье мы обсудим основные этапы, позволяющие отладку.</span><span class="sxs-lookup"><span data-stu-id="9134a-108">In this article, we discuss the key stages to enable debugging.</span></span>

- [<span data-ttu-id="9134a-109">Пометить надстройку для отладки</span><span class="sxs-lookup"><span data-stu-id="9134a-109">Mark the add-in for debugging</span></span>](#mark-your-add-in-for-debugging)
- [<span data-ttu-id="9134a-110">Настройка Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="9134a-110">Configure Visual Studio Code</span></span>](#configure-visual-studio-code)
- [<span data-ttu-id="9134a-111">Прикрепить Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="9134a-111">Attach Visual Studio Code</span></span>](#attach-visual-studio-code)
- [<span data-ttu-id="9134a-112">Debug</span><span class="sxs-lookup"><span data-stu-id="9134a-112">Debug</span></span>](#debug)

<span data-ttu-id="9134a-113">У вас есть несколько вариантов создания проекта надстройки.</span><span class="sxs-lookup"><span data-stu-id="9134a-113">You have several options for creating your add-in project.</span></span> <span data-ttu-id="9134a-114">В зависимости от используемого варианта действия могут отличаться.</span><span class="sxs-lookup"><span data-stu-id="9134a-114">Depending on the option you're using, the steps may vary.</span></span> <span data-ttu-id="9134a-115">Если вы использовали генератор Yeoman для Office надстроек для создания проекта надстройки (например, с помощью погона активации на основе [событий),](autolaunch.md)выполните  действия **yo office,** в противном случае выполните другие действия.</span><span class="sxs-lookup"><span data-stu-id="9134a-115">Where this is the case, if you used the Yeoman generator for Office Add-ins to create your add-in project (for example, by doing the [event-based activation walkthrough](autolaunch.md)), then follow the **yo office** steps, otherwise follow the **Other** steps.</span></span> <span data-ttu-id="9134a-116">Visual Studio Code должна быть по крайней мере версия 1.56.1.</span><span class="sxs-lookup"><span data-stu-id="9134a-116">Visual Studio Code should be at least version 1.56.1.</span></span>

## <a name="preview-debugging-for-the-event-based-activation-feature"></a><span data-ttu-id="9134a-117">Предварительная отладка функции активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="9134a-117">Preview debugging for the event-based activation feature</span></span>

<span data-ttu-id="9134a-118">Мы приглашаем вас попробовать возможности отладки для функции активации на основе событий!</span><span class="sxs-lookup"><span data-stu-id="9134a-118">We invite you to try out the debugging capability for the event-based activation feature!</span></span> <span data-ttu-id="9134a-119">Дайте нам знать о ваших сценариях и о том, как мы можем улучшить ситуацию, GitHub с помощью GitHub (см. раздел **Обратная** связь в конце этой страницы).</span><span class="sxs-lookup"><span data-stu-id="9134a-119">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="9134a-120">Чтобы просмотреть эту возможность Outlook на Windows, минимальная требуемая сборка составляет 16.0.13729.20000.</span><span class="sxs-lookup"><span data-stu-id="9134a-120">To preview this capability for Outlook on Windows, the minimum required build is 16.0.13729.20000.</span></span> <span data-ttu-id="9134a-121">Чтобы получить доступ Office бета-версий, присоединитесь к [программе Office Insider.](https://insider.office.com)</span><span class="sxs-lookup"><span data-stu-id="9134a-121">For access to Office beta builds, join the [Office Insider program](https://insider.office.com).</span></span>

## <a name="mark-your-add-in-for-debugging"></a><span data-ttu-id="9134a-122">Пометить надстройку для отладки</span><span class="sxs-lookup"><span data-stu-id="9134a-122">Mark your add-in for debugging</span></span>

1. <span data-ttu-id="9134a-123">Установите ключ `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` реестра.</span><span class="sxs-lookup"><span data-stu-id="9134a-123">Set the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.</span></span> <span data-ttu-id="9134a-124">`[Add-in ID]` является **Id** в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="9134a-124">`[Add-in ID]` is the **Id** in the add-in manifest.</span></span>

    <span data-ttu-id="9134a-125">**Yo office.** В окне командной строки перейдите к корневой папке надстройки и запустите следующую команду.</span><span class="sxs-lookup"><span data-stu-id="9134a-125">**yo office**: In a command line window, navigate to the root of your add-in folder then run the following command.</span></span>

    ```command&nbsp;line
    npm start
    ```

    <span data-ttu-id="9134a-126">В дополнение к построению кода и запуску локального сервера эта команда должна установить ключ реестра для этой `UseDirectDebugger` надстройки. `1`</span><span class="sxs-lookup"><span data-stu-id="9134a-126">In addition to building the code and starting the local server, this command should set the `UseDirectDebugger` registry key for this add-in to `1`.</span></span>

    <span data-ttu-id="9134a-127">**Другие:** Добавьте `UseDirectDebugger` ключ реестра под `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\` .</span><span class="sxs-lookup"><span data-stu-id="9134a-127">**Other**: Add the `UseDirectDebugger` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\`.</span></span> <span data-ttu-id="9134a-128">`[Add-in ID]`Замените **id из** манифеста надстройки.</span><span class="sxs-lookup"><span data-stu-id="9134a-128">Replace `[Add-in ID]` with the **Id** from the add-in manifest.</span></span> <span data-ttu-id="9134a-129">Установите ключ `1` реестра.</span><span class="sxs-lookup"><span data-stu-id="9134a-129">Set the registry key to `1`.</span></span>

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. <span data-ttu-id="9134a-130">Запустите Outlook (или перезапустите Outlook, если он уже открыт).</span><span class="sxs-lookup"><span data-stu-id="9134a-130">Start Outlook desktop (or restart Outlook if it's already open).</span></span>
1. <span data-ttu-id="9134a-131">Составить новое сообщение или назначение.</span><span class="sxs-lookup"><span data-stu-id="9134a-131">Compose a new message or appointment.</span></span> <span data-ttu-id="9134a-132">Вы должны увидеть следующий диалог.</span><span class="sxs-lookup"><span data-stu-id="9134a-132">You should see the following dialog.</span></span> <span data-ttu-id="9134a-133">Пока *не* взаимодействуйте с диалогом.</span><span class="sxs-lookup"><span data-stu-id="9134a-133">Do *not* interact with the dialog yet.</span></span>

    ![Снимок экрана диалогера обработки событий на основе отладки](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a><span data-ttu-id="9134a-135">Настройка Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="9134a-135">Configure Visual Studio Code</span></span>

### <a name="yo-office"></a><span data-ttu-id="9134a-136">yo office</span><span class="sxs-lookup"><span data-stu-id="9134a-136">yo office</span></span>

1. <span data-ttu-id="9134a-137">В окне командной строки откройте Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="9134a-137">Back in the command line window, open Visual Studio Code.</span></span>

    ```command&nbsp;line
    code .
    ```

1. <span data-ttu-id="9134a-138">В Visual Studio Code откройте файл **./.vscode/launch.js** и добавьте следующий отрывок в список конфигураций.</span><span class="sxs-lookup"><span data-stu-id="9134a-138">In Visual Studio Code, open the file **./.vscode/launch.json** and add the following excerpt to your list of configurations.</span></span> <span data-ttu-id="9134a-139">Сохраните изменения.</span><span class="sxs-lookup"><span data-stu-id="9134a-139">Save your changes.</span></span>

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

### <a name="other"></a><span data-ttu-id="9134a-140">Другое</span><span class="sxs-lookup"><span data-stu-id="9134a-140">Other</span></span>

1. <span data-ttu-id="9134a-141">Создайте новую папку под названием **Отладка** (возможно, в **папке Desktop).**</span><span class="sxs-lookup"><span data-stu-id="9134a-141">Create a new folder called **Debugging** (perhaps in your **Desktop** folder).</span></span>
1. <span data-ttu-id="9134a-142">Откройте Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="9134a-142">Open Visual Studio Code.</span></span>
1. <span data-ttu-id="9134a-143">Перейдите **к**  >  **открытой папке File Open,** перейдите к только что созданной папке, а затем выберите **Выберите папку**.</span><span class="sxs-lookup"><span data-stu-id="9134a-143">Go to **File** > **Open Folder**, navigate to the folder you just created, then choose **Select Folder**.</span></span>
1. <span data-ttu-id="9134a-144">В панели Действия выберите элемент **Отлаговка** (Ctrl+Shift+D).</span><span class="sxs-lookup"><span data-stu-id="9134a-144">On the Activity Bar, select the **Debug** item (Ctrl+Shift+D).</span></span>

    ![Снимок экрана значка Отлаговка в панели действий](../images/vs-code-debug.png)

1. <span data-ttu-id="9134a-146">Выберите **создание launch.jsссылки на файл.**</span><span class="sxs-lookup"><span data-stu-id="9134a-146">Select the **create a launch.json file** link.</span></span>

    ![Снимок экрана ссылки для создания launch.jsфайла в Visual Studio Code](../images/vs-code-create-launch.json.png)

1. <span data-ttu-id="9134a-148">В **отсеве Выберите среду** выберите **Edge: Запуск** для создания launch.jsфайла.</span><span class="sxs-lookup"><span data-stu-id="9134a-148">In the **Select Environment** dropdown, select **Edge: Launch** to create a launch.json file.</span></span>
1. <span data-ttu-id="9134a-149">Добавьте следующий отрывок в список конфигураций.</span><span class="sxs-lookup"><span data-stu-id="9134a-149">Add the following excerpt to your list of configurations.</span></span> <span data-ttu-id="9134a-150">Сохраните изменения.</span><span class="sxs-lookup"><span data-stu-id="9134a-150">Save your changes.</span></span>

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

## <a name="attach-visual-studio-code"></a><span data-ttu-id="9134a-151">Прикрепить Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="9134a-151">Attach Visual Studio Code</span></span>

1. <span data-ttu-id="9134a-152">Чтобы найтиbundle.jsнадстройки, \*\*\*\* откройте следующую папку в Windows Explorer и найдите **id** надстройки (найден в манифесте).</span><span class="sxs-lookup"><span data-stu-id="9134a-152">To find the add-in's **bundle.js**, open the following folder in Windows Explorer and search for your add-in's **Id** (found in the manifest).</span></span>

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    <span data-ttu-id="9134a-153">Откройте префикс папки с этим ID и скопируйте ее полный путь.</span><span class="sxs-lookup"><span data-stu-id="9134a-153">Open the folder prefixed with this ID and copy its full path.</span></span> <span data-ttu-id="9134a-154">В Visual Studio Code откройте **bundle.js** из этой папки.</span><span class="sxs-lookup"><span data-stu-id="9134a-154">In Visual Studio Code, open **bundle.js** from that folder.</span></span> <span data-ttu-id="9134a-155">Шаблон пути файла должен быть следующим:</span><span class="sxs-lookup"><span data-stu-id="9134a-155">The pattern of the file path should be as follows:</span></span>

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. <span data-ttu-id="9134a-156">Размыть точки bundle.js, где нужно остановить отладка.</span><span class="sxs-lookup"><span data-stu-id="9134a-156">Place breakpoints in bundle.js where you want the debugger to stop.</span></span>
1. <span data-ttu-id="9134a-157">В **отсеве DEBUG** выберите имя **Direct Debugging**, а затем выберите **Выполнить**.</span><span class="sxs-lookup"><span data-stu-id="9134a-157">In the **DEBUG** dropdown, select the name **Direct Debugging**, then select **Run**.</span></span>

    ![Снимок экрана выбора прямого отладки из параметров конфигурации в Visual Studio Code отладки](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a><span data-ttu-id="9134a-159">Debug</span><span class="sxs-lookup"><span data-stu-id="9134a-159">Debug</span></span>

1. <span data-ttu-id="9134a-160">После подтверждения того, что отладка присоединена, вернись в  Outlook и в диалоговом окне обработник на основе событий отладки выберите **ОК** .</span><span class="sxs-lookup"><span data-stu-id="9134a-160">After confirming that the debugger is attached, return to Outlook, and in the **Debug Event-based handler** dialog, choose **OK** .</span></span>

1. <span data-ttu-id="9134a-161">Теперь вы можете поразить точки Visual Studio Code, что позволит отключить код активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="9134a-161">You can now hit your breakpoints in Visual Studio Code, enabling you to debug your event-based activation code.</span></span>

## <a name="stop-debugging"></a><span data-ttu-id="9134a-162">Остановка отладки</span><span class="sxs-lookup"><span data-stu-id="9134a-162">Stop debugging</span></span>

<span data-ttu-id="9134a-163">Чтобы остановить отладку для остальной части текущего сеанса  Outlook рабочего стола, в диалоговом оклините обработите для отладки событий выберите **Отмена**.</span><span class="sxs-lookup"><span data-stu-id="9134a-163">To stop debugging for the rest of the current Outlook desktop session, in the **Debug Event-based handler** dialog, choose **Cancel**.</span></span> <span data-ttu-id="9134a-164">Чтобы повторно включить отладку, перезапустите Outlook рабочего стола.</span><span class="sxs-lookup"><span data-stu-id="9134a-164">To re-enable debugging, restart Outlook desktop.</span></span>

<span data-ttu-id="9134a-165">Чтобы предотвратить  отладку диалогового обработика событий на основе отладки и остановить отладку для последующих сеансов Outlook, удалите связанный ключ реестра или установите его `0` значение: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` .</span><span class="sxs-lookup"><span data-stu-id="9134a-165">To prevent the **Debug Event-based handler** dialog from popping up and stop debugging for subsequent Outlook sessions, delete the associated registry key or set its value to `0`: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.</span></span>

## <a name="see-also"></a><span data-ttu-id="9134a-166">См. также</span><span class="sxs-lookup"><span data-stu-id="9134a-166">See also</span></span>

- [<span data-ttu-id="9134a-167">Настройка надстройки Outlook для активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="9134a-167">Configure your Outlook add-in for event-based activation</span></span>](autolaunch.md)
- [<span data-ttu-id="9134a-168">Отладка надстройки с помощью журнала среды выполнения</span><span class="sxs-lookup"><span data-stu-id="9134a-168">Debug your add-in with runtime logging</span></span>](../testing/runtime-logging.md#runtime-logging-on-windows)
