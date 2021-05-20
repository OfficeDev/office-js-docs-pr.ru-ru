---
title: Отладить данные на основе Outlook (предварительный просмотр)
description: Узнайте, как отладить Outlook, которое реализует активацию на основе событий.
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
# <a name="debug-your-event-based-outlook-add-in-preview"></a><span data-ttu-id="99089-103">Отладить данные на основе Outlook (предварительный просмотр)</span><span class="sxs-lookup"><span data-stu-id="99089-103">Debug your event-based Outlook add-in (preview)</span></span>

<span data-ttu-id="99089-104">В этой статье содержатся рекомендации по отладке [при реализации активации событий](autolaunch.md) в надстройке.</span><span class="sxs-lookup"><span data-stu-id="99089-104">This article provides debugging guidance as you implement [event-based activation](autolaunch.md) in your add-in.</span></span> <span data-ttu-id="99089-105">Функция активации на основе событий в настоящее время находится в предварительном просмотре.</span><span class="sxs-lookup"><span data-stu-id="99089-105">The event-based activation feature is currently in preview.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="99089-106">Эта возможность отладки поддерживается только для предварительного просмотра Outlook в Windows с Microsoft 365 подпиской.</span><span class="sxs-lookup"><span data-stu-id="99089-106">This debugging capability is only supported for preview in Outlook on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="99089-107">Для получения дополнительной информации в [этой статье можно увидеть отладку Preview для раздела функции активации](#preview-debugging-for-the-event-based-activation-feature) на основе событий.</span><span class="sxs-lookup"><span data-stu-id="99089-107">For more information, see the [Preview debugging for the event-based activation feature](#preview-debugging-for-the-event-based-activation-feature) section in this article.</span></span>

<span data-ttu-id="99089-108">В этой статье мы обсуждаем ключевые этапы, позволяющие отладку.</span><span class="sxs-lookup"><span data-stu-id="99089-108">In this article, we discuss the key stages to enable debugging.</span></span>

- [<span data-ttu-id="99089-109">Отметь надстройку для отладки</span><span class="sxs-lookup"><span data-stu-id="99089-109">Mark the add-in for debugging</span></span>](#mark-your-add-in-for-debugging)
- [<span data-ttu-id="99089-110">Настройка Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="99089-110">Configure Visual Studio Code</span></span>](#configure-visual-studio-code)
- [<span data-ttu-id="99089-111">Прикрепите Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="99089-111">Attach Visual Studio Code</span></span>](#attach-visual-studio-code)
- [<span data-ttu-id="99089-112">Debug</span><span class="sxs-lookup"><span data-stu-id="99089-112">Debug</span></span>](#debug)

<span data-ttu-id="99089-113">У вас есть несколько вариантов создания надстройного проекта.</span><span class="sxs-lookup"><span data-stu-id="99089-113">You have several options for creating your add-in project.</span></span> <span data-ttu-id="99089-114">В зависимости от используемого варианта шаги могут отличаться.</span><span class="sxs-lookup"><span data-stu-id="99089-114">Depending on the option you're using, the steps may vary.</span></span> <span data-ttu-id="99089-115">В этом случае, если вы использовали генератор Yeoman для Office дополнительных виленок для создания надстройки проекта (например, делая [пошаговое руководство активации на основе событий),](autolaunch.md)то следуйте **шагам офиса yo,** в противном случае **следуйте другим** шагам.</span><span class="sxs-lookup"><span data-stu-id="99089-115">Where this is the case, if you used the Yeoman generator for Office Add-ins to create your add-in project (for example, by doing the [event-based activation walkthrough](autolaunch.md)), then follow the **yo office** steps, otherwise follow the **Other** steps.</span></span> <span data-ttu-id="99089-116">Visual Studio Code должна быть по крайней мере версия 1.56.1.</span><span class="sxs-lookup"><span data-stu-id="99089-116">Visual Studio Code should be at least version 1.56.1.</span></span>

## <a name="preview-debugging-for-the-event-based-activation-feature"></a><span data-ttu-id="99089-117">Предварительная отладка функции активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="99089-117">Preview debugging for the event-based activation feature</span></span>

<span data-ttu-id="99089-118">Мы приглашаем Вас опробовать возможности отладки для функции активации на основе событий!</span><span class="sxs-lookup"><span data-stu-id="99089-118">We invite you to try out the debugging capability for the event-based activation feature!</span></span> <span data-ttu-id="99089-119">Сообщите нам о ваших сценариях и о том, как мы можем улучшить их, дав нам обратную связь GitHub **(см.** раздел Обратная связь в конце этой страницы).</span><span class="sxs-lookup"><span data-stu-id="99089-119">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="99089-120">Для просмотра этой возможности Outlook на Windows, минимальная требуемая сборка составляет 16.0.13729.20000.</span><span class="sxs-lookup"><span data-stu-id="99089-120">To preview this capability for Outlook on Windows, the minimum required build is 16.0.13729.20000.</span></span> <span data-ttu-id="99089-121">Для доступа к Office бета-сборки, присоединяйтесь [к Office Insider](https://insider.office.com).</span><span class="sxs-lookup"><span data-stu-id="99089-121">For access to Office beta builds, join the [Office Insider program](https://insider.office.com).</span></span>

## <a name="mark-your-add-in-for-debugging"></a><span data-ttu-id="99089-122">Отметите надстройку для отладки</span><span class="sxs-lookup"><span data-stu-id="99089-122">Mark your add-in for debugging</span></span>

1. <span data-ttu-id="99089-123">Установите ключ реестра `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` .</span><span class="sxs-lookup"><span data-stu-id="99089-123">Set the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.</span></span> <span data-ttu-id="99089-124">`[Add-in ID]` является **идентификатором** в надстройной манифесте.</span><span class="sxs-lookup"><span data-stu-id="99089-124">`[Add-in ID]` is the **Id** in the add-in manifest.</span></span>

    <span data-ttu-id="99089-125">**yo office**: В окне командной строки перейдите к корню папки с надстройки, а затем запустите следующую команду.</span><span class="sxs-lookup"><span data-stu-id="99089-125">**yo office**: In a command line window, navigate to the root of your add-in folder then run the following command.</span></span>

    ```command&nbsp;line
    npm start
    ```

    <span data-ttu-id="99089-126">В дополнение к созданию кода и запуску локального сервера, эта команда должна `UseDirectDebugger` установить ключ реестра для этого дополнения к `1` .</span><span class="sxs-lookup"><span data-stu-id="99089-126">In addition to building the code and starting the local server, this command should set the `UseDirectDebugger` registry key for this add-in to `1`.</span></span>

    <span data-ttu-id="99089-127">**Другие**: Добавить `UseDirectDebugger` ключ реестра под `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\` .</span><span class="sxs-lookup"><span data-stu-id="99089-127">**Other**: Add the `UseDirectDebugger` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\`.</span></span> <span data-ttu-id="99089-128">`[Add-in ID]`Замените **id** из надстройного манифеста.</span><span class="sxs-lookup"><span data-stu-id="99089-128">Replace `[Add-in ID]` with the **Id** from the add-in manifest.</span></span> <span data-ttu-id="99089-129">Установите ключ реестра к `1` .</span><span class="sxs-lookup"><span data-stu-id="99089-129">Set the registry key to `1`.</span></span>

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. <span data-ttu-id="99089-130">Начните Outlook (или перезапустите Outlook, если он уже открыт).</span><span class="sxs-lookup"><span data-stu-id="99089-130">Start Outlook desktop (or restart Outlook if it's already open).</span></span>
1. <span data-ttu-id="99089-131">Составьте новое сообщение или встречу.</span><span class="sxs-lookup"><span data-stu-id="99089-131">Compose a new message or appointment.</span></span> <span data-ttu-id="99089-132">Вы должны увидеть следующий диалог.</span><span class="sxs-lookup"><span data-stu-id="99089-132">You should see the following dialog.</span></span> <span data-ttu-id="99089-133">Пока *не* взаимодействуйте с диалогом.</span><span class="sxs-lookup"><span data-stu-id="99089-133">Do *not* interact with the dialog yet.</span></span>

    ![Скриншот диалога обработчика на основе событий Отбага](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a><span data-ttu-id="99089-135">Настройка Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="99089-135">Configure Visual Studio Code</span></span>

### <a name="yo-office"></a><span data-ttu-id="99089-136">йо-офис</span><span class="sxs-lookup"><span data-stu-id="99089-136">yo office</span></span>

1. <span data-ttu-id="99089-137">Вернуться в окно командной строки, открыть Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="99089-137">Back in the command line window, open Visual Studio Code.</span></span>

    ```command&nbsp;line
    code .
    ```

1. <span data-ttu-id="99089-138">В Visual Studio Code откройте файл **./.vscode/launch.jsи добавьте** следующий отрывок в список конфигураций.</span><span class="sxs-lookup"><span data-stu-id="99089-138">In Visual Studio Code, open the file **./.vscode/launch.json** and add the following excerpt to your list of configurations.</span></span> <span data-ttu-id="99089-139">Сохраните изменения.</span><span class="sxs-lookup"><span data-stu-id="99089-139">Save your changes.</span></span>

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

### <a name="other"></a><span data-ttu-id="99089-140">Другое</span><span class="sxs-lookup"><span data-stu-id="99089-140">Other</span></span>

1. <span data-ttu-id="99089-141">Создайте новую папку под **названием Debugging** (возможно, в **папке Desktop).**</span><span class="sxs-lookup"><span data-stu-id="99089-141">Create a new folder called **Debugging** (perhaps in your **Desktop** folder).</span></span>
1. <span data-ttu-id="99089-142">Открытая Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="99089-142">Open Visual Studio Code.</span></span>
1. <span data-ttu-id="99089-143">Перейдите в **File**  >  **Open Folder,** перейдите к папке, которую вы только что создали, а затем **выберите Select Folder**.</span><span class="sxs-lookup"><span data-stu-id="99089-143">Go to **File** > **Open Folder**, navigate to the folder you just created, then choose **Select Folder**.</span></span>
1. <span data-ttu-id="99089-144">В баре активности выберите **элемент отладки** (Ctrl-Shift-D).</span><span class="sxs-lookup"><span data-stu-id="99089-144">On the Activity Bar, select the **Debug** item (Ctrl+Shift+D).</span></span>

    ![Скриншот значка отладки на баре активности](../images/vs-code-debug.png)

1. <span data-ttu-id="99089-146">Выберите **создать launch.jsссылке файла.**</span><span class="sxs-lookup"><span data-stu-id="99089-146">Select the **create a launch.json file** link.</span></span>

    ![Скриншот ссылки для создания launch.jsфайла в Visual Studio Code](../images/vs-code-create-launch.json.png)

1. <span data-ttu-id="99089-148">В **выберите среду** падения, выберите **Край: Запуск для** создания launch.jsв файле.</span><span class="sxs-lookup"><span data-stu-id="99089-148">In the **Select Environment** dropdown, select **Edge: Launch** to create a launch.json file.</span></span>
1. <span data-ttu-id="99089-149">Добавьте следующий отрывок в список конфигураций.</span><span class="sxs-lookup"><span data-stu-id="99089-149">Add the following excerpt to your list of configurations.</span></span> <span data-ttu-id="99089-150">Сохраните изменения.</span><span class="sxs-lookup"><span data-stu-id="99089-150">Save your changes.</span></span>

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

## <a name="attach-visual-studio-code"></a><span data-ttu-id="99089-151">Прикрепите Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="99089-151">Attach Visual Studio Code</span></span>

1. <span data-ttu-id="99089-152">Чтобы найти дополнительные данные **bundle.js,** откройте следующую папку в Windows Explorer и ищите **идентификатор надстройки** (найдено в манифесте).</span><span class="sxs-lookup"><span data-stu-id="99089-152">To find the add-in's **bundle.js**, open the following folder in Windows Explorer and search for your add-in's **Id** (found in the manifest).</span></span>

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    <span data-ttu-id="99089-153">Откройте папку, наклееную на этот идентификатор, и скопировать ее полный путь.</span><span class="sxs-lookup"><span data-stu-id="99089-153">Open the folder prefixed with this ID and copy its full path.</span></span> <span data-ttu-id="99089-154">В Visual Studio Code, откройте **bundle.jsиз** этой папки.</span><span class="sxs-lookup"><span data-stu-id="99089-154">In Visual Studio Code, open **bundle.js** from that folder.</span></span> <span data-ttu-id="99089-155">Шаблон пути файла должен быть следующим:</span><span class="sxs-lookup"><span data-stu-id="99089-155">The pattern of the file path should be as follows:</span></span>

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. <span data-ttu-id="99089-156">Поместите точки разрыва в bundle.js где вы хотите, чтобы отбегер, чтобы остановить.</span><span class="sxs-lookup"><span data-stu-id="99089-156">Place breakpoints in bundle.js where you want the debugger to stop.</span></span>
1. <span data-ttu-id="99089-157">В **отключке DEBUG** выберите имя **Прямая отладка,** а затем выберите **Run.**</span><span class="sxs-lookup"><span data-stu-id="99089-157">In the **DEBUG** dropdown, select the name **Direct Debugging**, then select **Run**.</span></span>

    ![Скриншот выбора прямого отладки из вариантов конфигурации в Visual Studio Code отладки](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a><span data-ttu-id="99089-159">Debug</span><span class="sxs-lookup"><span data-stu-id="99089-159">Debug</span></span>

1. <span data-ttu-id="99089-160">После подтверждения того, что отладка прилагается, вернитесь к Outlook, и в **диалоге обработчика на основе события отладки** выберите **OK** .</span><span class="sxs-lookup"><span data-stu-id="99089-160">After confirming that the debugger is attached, return to Outlook, and in the **Debug Event-based handler** dialog, choose **OK** .</span></span>

1. <span data-ttu-id="99089-161">Теперь вы можете поразить точки разрыва в Visual Studio Code, что позволит вам отладить код активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="99089-161">You can now hit your breakpoints in Visual Studio Code, enabling you to debug your event-based activation code.</span></span>

## <a name="stop-debugging"></a><span data-ttu-id="99089-162">Остановить отладку</span><span class="sxs-lookup"><span data-stu-id="99089-162">Stop debugging</span></span>

<span data-ttu-id="99089-163">Чтобы остановить отладку для остальной части текущего сеанса Outlook рабочего стола, в **диалоге обработчиков на основе события отладки,** выберите **Отменить**.</span><span class="sxs-lookup"><span data-stu-id="99089-163">To stop debugging for the rest of the current Outlook desktop session, in the **Debug Event-based handler** dialog, choose **Cancel**.</span></span> <span data-ttu-id="99089-164">Чтобы повторно включить отладку, перезапустите Outlook столе.</span><span class="sxs-lookup"><span data-stu-id="99089-164">To re-enable debugging, restart Outlook desktop.</span></span>

<span data-ttu-id="99089-165">Чтобы предотвратить **выскакивание диалога обработчика на основе событий Отладки** и прекратить отладку для последующих сеансов Outlook, удалите связанный ключ реестра или установите его `0` значение: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` .</span><span class="sxs-lookup"><span data-stu-id="99089-165">To prevent the **Debug Event-based handler** dialog from popping up and stop debugging for subsequent Outlook sessions, delete the associated registry key or set its value to `0`: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.</span></span>

## <a name="see-also"></a><span data-ttu-id="99089-166">См. также</span><span class="sxs-lookup"><span data-stu-id="99089-166">See also</span></span>

- [<span data-ttu-id="99089-167">Настройте Outlook для активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="99089-167">Configure your Outlook add-in for event-based activation</span></span>](autolaunch.md)
- [<span data-ttu-id="99089-168">Отладка надстройки с помощью журнала среды выполнения</span><span class="sxs-lookup"><span data-stu-id="99089-168">Debug your add-in with runtime logging</span></span>](../testing/runtime-logging.md#runtime-logging-on-windows)
