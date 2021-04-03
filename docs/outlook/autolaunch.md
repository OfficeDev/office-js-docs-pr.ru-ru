---
title: Настройка надстройки Outlook для активации на основе событий (предварительный просмотр)
description: Узнайте, как настроить надстройку Outlook для активации на основе событий.
ms.topic: article
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: b9a4460b05af14f57eecb1bf4181c706843537b2
ms.sourcegitcommit: 074526a6dca8381dbdabf2705474c5ae6753b829
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/02/2021
ms.locfileid: "51506149"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="1946c-103">Настройка надстройки Outlook для активации на основе событий (предварительный просмотр)</span><span class="sxs-lookup"><span data-stu-id="1946c-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="1946c-104">Без функции активации на основе событий пользователю необходимо явно запустить надстройки для выполнения задач.</span><span class="sxs-lookup"><span data-stu-id="1946c-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="1946c-105">Эта функция позволяет надстройки выполнять задачи на основе определенных событий, особенно для операций, применимых к каждому элементу.</span><span class="sxs-lookup"><span data-stu-id="1946c-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="1946c-106">Вы также можете интегрироваться с области задач и функциональными возможностями без пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="1946c-106">You can also integrate with the task pane and UI-less functionality.</span></span> <span data-ttu-id="1946c-107">В настоящее время поддерживаются следующие события.</span><span class="sxs-lookup"><span data-stu-id="1946c-107">At present, the following events are supported.</span></span>

- <span data-ttu-id="1946c-108">`OnNewMessageCompose`: При сочинении нового сообщения (включает ответ, ответ все и вперед)</span><span class="sxs-lookup"><span data-stu-id="1946c-108">`OnNewMessageCompose`: On composing a new message (includes reply, reply all, and forward)</span></span>
- <span data-ttu-id="1946c-109">`OnNewAppointmentOrganizer`: При создании нового назначения</span><span class="sxs-lookup"><span data-stu-id="1946c-109">`OnNewAppointmentOrganizer`: On creating a new appointment</span></span>

  > [!IMPORTANT]
  > <span data-ttu-id="1946c-110">Эта функция **не активируется** при редактировании элемента, например черновика или существующей встречи.</span><span class="sxs-lookup"><span data-stu-id="1946c-110">This feature does **not** activate on editing an item, for example, a draft or an existing appointment.</span></span>

<span data-ttu-id="1946c-111">К концу этого погона у вас будет надстройка, которая запускается всякий раз, когда создается новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="1946c-111">By the end of this walkthrough, you'll have an add-in that runs whenever a new message is created.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1946c-112">Эта функция поддерживается только [для](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) предварительного просмотра в Outlook в Интернете и в Windows с подпиской microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="1946c-112">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="1946c-113">Дополнительные [сведения см. в](#how-to-preview-the-event-based-activation-feature) статье Как просмотреть функцию активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="1946c-113">See [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="1946c-114">Так как функции предварительного просмотра могут изменяться без предварительного уведомления, их не следует использовать в надстройки производства.</span><span class="sxs-lookup"><span data-stu-id="1946c-114">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="1946c-115">Просмотр функции активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="1946c-115">How to preview the event-based activation feature</span></span>

<span data-ttu-id="1946c-116">Мы приглашаем вас попробовать функцию активации на основе событий!</span><span class="sxs-lookup"><span data-stu-id="1946c-116">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="1946c-117">Дайте нам знать ваши сценарии и как мы можем улучшить, предоставляя нам отзывы через GitHub (см. раздел **Обратная** связь в конце этой страницы).</span><span class="sxs-lookup"><span data-stu-id="1946c-117">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="1946c-118">Чтобы просмотреть эту функцию:</span><span class="sxs-lookup"><span data-stu-id="1946c-118">To preview this feature:</span></span>

- <span data-ttu-id="1946c-119">Для Outlook в Интернете:</span><span class="sxs-lookup"><span data-stu-id="1946c-119">For Outlook on the web:</span></span>
  - <span data-ttu-id="1946c-120">[Настройка целевого выпуска для клиента Microsoft 365.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="1946c-120">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="1946c-121">Ссылка на **бета-библиотеку** на CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="1946c-121">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="1946c-122">Файл [определения типа](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) для компиляции и IntelliSense typeScript находится в CDN и [DefinitelyTyped.](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)</span><span class="sxs-lookup"><span data-stu-id="1946c-122">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="1946c-123">Эти типы можно установить с `npm install --save-dev @types/office-js-preview` помощью .</span><span class="sxs-lookup"><span data-stu-id="1946c-123">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="1946c-124">Для Outlook on Windows: минимальная требуемая сборка — 16.0.13729.20000.</span><span class="sxs-lookup"><span data-stu-id="1946c-124">For Outlook on Windows: The minimum required build is 16.0.13729.20000.</span></span> <span data-ttu-id="1946c-125">Присоединяйтесь [к программе инсайдерской службы](https://insider.office.com) Office для доступа к бета-сборкам Office.</span><span class="sxs-lookup"><span data-stu-id="1946c-125">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="1946c-126">Настройка среды</span><span class="sxs-lookup"><span data-stu-id="1946c-126">Set up your environment</span></span>

<span data-ttu-id="1946c-127">Выполните [быстрый запуск Outlook,](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) который создает проект надстройки с генератором Yeoman для надстройок Office.</span><span class="sxs-lookup"><span data-stu-id="1946c-127">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="1946c-128">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="1946c-128">Configure the manifest</span></span>

<span data-ttu-id="1946c-129">Чтобы включить активацию надстройки на основе событий, необходимо настроить элемент [Runtimes](../reference/manifest/runtimes.md) и точку расширения [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) в узле `VersionOverridesV1_1` манифеста.</span><span class="sxs-lookup"><span data-stu-id="1946c-129">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="1946c-130">Пока это `DesktopFormFactor` единственный поддерживаемый форм-фактор.</span><span class="sxs-lookup"><span data-stu-id="1946c-130">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="1946c-131">В редакторе кода откройте проект быстрого запуска.</span><span class="sxs-lookup"><span data-stu-id="1946c-131">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="1946c-132">Откройте **manifest.xml** файл, расположенный в корне проекта.</span><span class="sxs-lookup"><span data-stu-id="1946c-132">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="1946c-133">Выберите весь узел (включая открытые и закрываемые теги) и замените его на следующий XML, а затем `<VersionOverrides>` сохраните изменения.</span><span class="sxs-lookup"><span data-stu-id="1946c-133">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.</span></span>

```XML
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- Event-based activation happens in a lightweight runtime.-->
        <Runtimes>
          <!-- HTML file including reference to or inline JavaScript event handlers.
               This is used by Outlook on the web. -->
          <Runtime resid="WebViewRuntime.Url">
            <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
            <Override type="javascript" resid="JSRuntime.Url"/>
          </Runtime>
        </Runtimes>
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgReadGroup">
                <Label resid="GroupLabel" />
                <Control xsi:type="Button" id="msgReadOpenPaneButton">
                  <Label resid="TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="TaskpaneButton.Label" />
                    <Description resid="TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16" />
                    <bt:Image size="32" resid="Icon.32x32" />
                    <bt:Image size="80" resid="Icon.80x80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
                <Control xsi:type="Button" id="ActionButton">
                  <Label resid="ActionButton.Label"/>
                  <Supertip>
                    <Title resid="ActionButton.Label"/>
                    <Description resid="ActionButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>action</FunctionName>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>

          <!-- Can configure other command surface extension points for add-in command support. -->

          <!-- Enable launching the add-in on the included events. -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
            </LaunchEvents>
            <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
            <SourceLocation resid="WebViewRuntime.Url"/>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html" />
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html" />
        <bt:Url id="WebViewRuntime.Url" DefaultValue="https://localhost:3000/commands.html" />
        <!-- Entry needed for Outlook Desktop. -->
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/src/commands/commands.js" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="Contoso Add-in"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Show Taskpane"/>
        <bt:String id="ActionButton.Label" DefaultValue="Perform an action"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Opens a pane displaying all available properties."/>
        <bt:String id="ActionButton.Tooltip" DefaultValue="Perform an action when clicked."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

<span data-ttu-id="1946c-134">Outlook в Windows использует файл JavaScript, в то время как Outlook в Интернете использует HTML-файл, который может ссылаться на тот же файл JavaScript.</span><span class="sxs-lookup"><span data-stu-id="1946c-134">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="1946c-135">Необходимо предоставить ссылки на оба этих файла в узле манифеста, так как платформа Outlook в конечном счете определяет, следует ли использовать HTML или JavaScript на основе `Resources` клиента Outlook.</span><span class="sxs-lookup"><span data-stu-id="1946c-135">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="1946c-136">Таким образом, чтобы настроить обработку событий, укадь расположение HTML в элементе, а затем в его детском элементе укаймляй расположение файла JavaScript, вписаного или ссылаемого `Runtime` `Override` HTML.</span><span class="sxs-lookup"><span data-stu-id="1946c-136">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="1946c-137">Дополнительные новости о манифестах надстройок Outlook см. в манифестах [надстройки Outlook.](manifests.md)</span><span class="sxs-lookup"><span data-stu-id="1946c-137">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="1946c-138">Реализация обработки событий</span><span class="sxs-lookup"><span data-stu-id="1946c-138">Implement event handling</span></span>

<span data-ttu-id="1946c-139">Для выбранных событий необходимо реализовать обработку.</span><span class="sxs-lookup"><span data-stu-id="1946c-139">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="1946c-140">В этом сценарии вы добавим обработку для составления новых элементов.</span><span class="sxs-lookup"><span data-stu-id="1946c-140">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="1946c-141">В том же проекте быстрого запуска откройте **файл ./src/commands/commands.js** в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="1946c-141">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="1946c-142">После `action` функции вставьте следующие функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="1946c-142">After the `action` function, insert the following JavaScript functions.</span></span>

    ```js
    function onMessageComposeHandler(event) {
      setSubject(event);
    }
    function onAppointmentComposeHandler(event) {
      setSubject(event);
    }
    function setSubject(event) {
      Office.context.mailbox.item.subject.setAsync(
        "Set by an event-based add-in!",
        {
          "asyncContext" : event
        },
        function (asyncResult) {
          // Handle success or error.
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
          }
    
          // Call event.completed() after all work is done.
          asyncResult.asyncContext.completed();
        });
    }
    ```

1. <span data-ttu-id="1946c-143">Чтобы функции работали в **Outlook** в Интернете с этим проектом, созданным генератором Yeoman для надстройок Office, добавьте следующие утверждения в конце файла.</span><span class="sxs-lookup"><span data-stu-id="1946c-143">For the functions to work in **Outlook on the web** with this project generated by the Yeoman generator for Office Add-ins, add the following statements at the end of the file.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

1. <span data-ttu-id="1946c-144">Чтобы функции работали в **Outlook on Windows,** добавьте следующий код JavaScript в конце файла.</span><span class="sxs-lookup"><span data-stu-id="1946c-144">For the functions to work in **Outlook on Windows**, add the following JavaScript code at the end of the file.</span></span>

    ```js
    if (Office.actions) {
      // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
      Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
      Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    }
    ```

    <span data-ttu-id="1946c-145">**Примечание.** Проверка гарантирует, что Outlook в Интернете `Office.actions` игнорирует эти утверждения.</span><span class="sxs-lookup"><span data-stu-id="1946c-145">**Note**: Checking for `Office.actions` ensures that Outlook on the web ignores these statements.</span></span>

1. <span data-ttu-id="1946c-146">Сохраните изменения.</span><span class="sxs-lookup"><span data-stu-id="1946c-146">Save your changes.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="1946c-147">Проверка</span><span class="sxs-lookup"><span data-stu-id="1946c-147">Try it out</span></span>

1. <span data-ttu-id="1946c-148">Выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="1946c-148">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="1946c-149">После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен) и будет загружена ваша неопубликованная надстройка.</span><span class="sxs-lookup"><span data-stu-id="1946c-149">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

1. <span data-ttu-id="1946c-150">Создайте новое сообщение в веб-версии Outlook.</span><span class="sxs-lookup"><span data-stu-id="1946c-150">In Outlook on the web, create a new message.</span></span>

    ![Снимок экрана окна сообщения в Outlook в Интернете с набором темы для составить](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="1946c-152">В Outlook on Windows создайте новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="1946c-152">In Outlook on Windows, create a new message.</span></span>

    ![Снимок экрана окна сообщения в Outlook on Windows с набором субъекта на композит](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > <span data-ttu-id="1946c-154">Если вы видите ошибку "Мы не можем открыть эту надстройку из localhost", необходимо включить освобождение от циклов.</span><span class="sxs-lookup"><span data-stu-id="1946c-154">If you see the error "We can't open this add-in from localhost," you'll need to enable a loopback exemption.</span></span>
    >
    > 1. <span data-ttu-id="1946c-155">Закройте Outlook.</span><span class="sxs-lookup"><span data-stu-id="1946c-155">Close Outlook.</span></span>
    > 2. <span data-ttu-id="1946c-156">Откройте диспетчер **задач** и убедитесь, что **msoadfs.exe** процесс не запущен.</span><span class="sxs-lookup"><span data-stu-id="1946c-156">Open the **Task Manager** and ensure that the **msoadfs.exe** process is not running.</span></span>
    > 3. <span data-ttu-id="1946c-157">Выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="1946c-157">Run the following command.</span></span>
    >
    >     ```command&nbsp;line
    >     call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >     ```
    >
    > 4. <span data-ttu-id="1946c-158">Перезапустите Outlook.</span><span class="sxs-lookup"><span data-stu-id="1946c-158">Restart Outlook.</span></span>

## <a name="debug"></a><span data-ttu-id="1946c-159">Debug</span><span class="sxs-lookup"><span data-stu-id="1946c-159">Debug</span></span>

<span data-ttu-id="1946c-160">При реализации собственных функций может потребоваться отламывка кода.</span><span class="sxs-lookup"><span data-stu-id="1946c-160">As you implement your own functionality, you may need to debug your code.</span></span> <span data-ttu-id="1946c-161">Инструкции по отламывить активацию надстройки на основе событий см. в инструкции по отламывить надстройки Outlook на основе [событий.](debug-autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="1946c-161">For guidance on how to debug event-based add-in activation, see [Debug your event-based Outlook add-in](debug-autolaunch.md).</span></span>

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="1946c-162">Поведение и ограничения активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="1946c-162">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="1946c-163">Надстройки, которые активируются на основе событий, как ожидается, будут короткими, легкими и максимально неинвазивными.</span><span class="sxs-lookup"><span data-stu-id="1946c-163">Add-ins that activate based on events are expected to be short-running, lightweight, and as non-invasive as possible.</span></span> <span data-ttu-id="1946c-164">Чтобы сигнализировать, что надстройка завершила обработку события запуска, рекомендуется использовать метод вызова `event.completed` надстройки.</span><span class="sxs-lookup"><span data-stu-id="1946c-164">To signal that your add-in has completed processing the launch event, we recommend you have your add-in call the `event.completed` method.</span></span> <span data-ttu-id="1946c-165">Если этот вызов не будет выполнен, надстройка будет работать в течение примерно 300 секунд, что является максимальным сроком, разрешенным для запуска надстроек на основе событий. Надстройка также заканчивается, когда пользователь закрывает окно записи.</span><span class="sxs-lookup"><span data-stu-id="1946c-165">If that call is not made, the add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="1946c-166">Если у пользователя есть несколько надстройок, которые подписаны на одно и то же событие, платформа Outlook запускает надстройки без определенного порядка.</span><span class="sxs-lookup"><span data-stu-id="1946c-166">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="1946c-167">В настоящее время можно активно запускать только пять надстройок на основе событий.</span><span class="sxs-lookup"><span data-stu-id="1946c-167">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="1946c-168">Все дополнительные надстройки отодвигаются в очередь, а затем запускаются по мере завершения или отключения ранее активных надстроек.</span><span class="sxs-lookup"><span data-stu-id="1946c-168">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="1946c-169">Пользователь может переключаться или перемещаться от текущего элемента почты, где надстройка начала работать.</span><span class="sxs-lookup"><span data-stu-id="1946c-169">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="1946c-170">Запущенная надстройка завершит свою работу в фоновом режиме.</span><span class="sxs-lookup"><span data-stu-id="1946c-170">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="1946c-171">Некоторые Office.js API, которые изменяют или изменяют пользовательский интерфейс, не допускаются из надстройок на основе событий. Следующие API заблокированы:</span><span class="sxs-lookup"><span data-stu-id="1946c-171">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="1946c-172">В `Office.context.auth` статье:</span><span class="sxs-lookup"><span data-stu-id="1946c-172">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`
- <span data-ttu-id="1946c-173">В `Office.context.mailbox` статье:</span><span class="sxs-lookup"><span data-stu-id="1946c-173">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="1946c-174">В `Office.context.mailbox.item` статье:</span><span class="sxs-lookup"><span data-stu-id="1946c-174">Under `Office.context.mailbox.item`:</span></span>
  - `close`
- <span data-ttu-id="1946c-175">В `Office.context.ui` статье:</span><span class="sxs-lookup"><span data-stu-id="1946c-175">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a><span data-ttu-id="1946c-176">См. также</span><span class="sxs-lookup"><span data-stu-id="1946c-176">See also</span></span>

<span data-ttu-id="1946c-177">[Манифесты надстройки](manifests.md) 
 Outlook [Отламывка](debug-autolaunch.md) надстроек на основе событий</span><span class="sxs-lookup"><span data-stu-id="1946c-177">[Outlook add-in manifests](manifests.md)
[How to debug event-based add-ins](debug-autolaunch.md)</span></span>
