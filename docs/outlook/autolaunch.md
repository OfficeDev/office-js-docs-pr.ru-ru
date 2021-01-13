---
title: Настройка надстройки Outlook для активации на основе событий (предварительная версия)
description: Узнайте, как настроить надстройку Outlook для активации на основе событий.
ms.topic: article
ms.date: 01/06/2021
localization_priority: Normal
ms.openlocfilehash: d6893733af52bba7917531b2e8d5a442ce3dcd77
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839833"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="9a8f1-103">Настройка надстройки Outlook для активации на основе событий (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="9a8f1-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="9a8f1-104">Без функции активации на основе событий пользователю необходимо явным образом запустить надстройки для выполнения своих задач.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="9a8f1-105">Эта функция позволяет надстройка выполнять задачи на основе определенных событий, особенно для операций, которые применяются к каждому элементу.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="9a8f1-106">Вы также можете интегрироваться с функциональностью области задач и без пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-106">You can also integrate with the task pane and UI-less functionality.</span></span> <span data-ttu-id="9a8f1-107">В настоящее время поддерживаются следующие события.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-107">At present, the following events are supported.</span></span>

- <span data-ttu-id="9a8f1-108">`OnNewMessageCompose`: при составления нового сообщения (включая ответ, ответ всем и переад. );</span><span class="sxs-lookup"><span data-stu-id="9a8f1-108">`OnNewMessageCompose`: On composing a new message (includes reply, reply all, and forward)</span></span>
- <span data-ttu-id="9a8f1-109">`OnNewAppointmentOrganizer`: При создании новой встречи</span><span class="sxs-lookup"><span data-stu-id="9a8f1-109">`OnNewAppointmentOrganizer`: On creating a new appointment</span></span>

  > [!IMPORTANT]
  > <span data-ttu-id="9a8f1-110">Эта функция не **активируется** при редактировании элемента, например черновика или существующей встречи.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-110">This feature does **not** activate on editing an item, for example, a draft or an existing appointment.</span></span>

<span data-ttu-id="9a8f1-111">К концу этого побока у вас будет надстройка, которая запускается при всех создания нового сообщения.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-111">By the end of this walkthrough, you'll have an add-in that runs whenever a new message is created.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9a8f1-112">Эта функция поддерживается только для [предварительного](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) просмотра в Outlook в Интернете с подпиской на Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-112">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web with a Microsoft 365 subscription.</span></span> <span data-ttu-id="9a8f1-113">Дополнительные [сведения см.](#how-to-preview-the-event-based-activation-feature) в этой статье, чтобы просмотреть функцию активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-113">See [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="9a8f1-114">Поскольку функции предварительного просмотра могут быть без предварительного уведомления, их не следует использовать в производственных надстройки.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-114">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="9a8f1-115">Просмотр функции активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="9a8f1-115">How to preview the event-based activation feature</span></span>

<span data-ttu-id="9a8f1-116">Мы предлагаем вам попробовать функцию активации на основе событий!</span><span class="sxs-lookup"><span data-stu-id="9a8f1-116">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="9a8f1-117">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span><span class="sxs-lookup"><span data-stu-id="9a8f1-117">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="9a8f1-118">Чтобы просмотреть эту функцию:</span><span class="sxs-lookup"><span data-stu-id="9a8f1-118">To preview this feature:</span></span>

- <span data-ttu-id="9a8f1-119">Ссылка на **бета-версию** библиотеки в CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="9a8f1-119">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="9a8f1-120">Файл [определения типа для](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) компиляции и IntelliSense TypeScript находится в CDN и [DefinitelyTyped.](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)</span><span class="sxs-lookup"><span data-stu-id="9a8f1-120">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="9a8f1-121">Эти типы можно установить с помощью `npm install --save-dev @types/office-js-preview` .</span><span class="sxs-lookup"><span data-stu-id="9a8f1-121">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="9a8f1-122">[Настройка целевого выпуска в клиенте Microsoft 365.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="9a8f1-122">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="9a8f1-123">Настройка среды</span><span class="sxs-lookup"><span data-stu-id="9a8f1-123">Set up your environment</span></span>

<span data-ttu-id="9a8f1-124">Завершите [краткое начало работы с Outlook,](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) которое создает проект надстройки с помощью генератора Yeoman для надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-124">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="9a8f1-125">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="9a8f1-125">Configure the manifest</span></span>

<span data-ttu-id="9a8f1-126">Чтобы включить активацию надстройки на основе событий, необходимо настроить элемент [Runtimes](../reference/manifest/runtimes.md) и точку расширения [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) в манифесте.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-126">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the manifest.</span></span> <span data-ttu-id="9a8f1-127">На данный момент `DesktopFormFactor` это единственный поддерживаемый форм-фактор.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-127">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="9a8f1-128">В редакторе кода откройте проект быстрого запуска.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-128">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="9a8f1-129">Откройте файл **manifest.xml,** расположенный в корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-129">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="9a8f1-130">Выберите весь узел (включая открытые и закрываемые `<VersionOverrides>` теги) и замените его на следующий XML-</span><span class="sxs-lookup"><span data-stu-id="9a8f1-130">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/runtime.js" />
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

<span data-ttu-id="9a8f1-131">Outlook для Windows использует файл JavaScript, а Outlook в Интернете использует HTML-файл, ссылаясь на тот же файл JavaScript.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-131">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that references the same JavaScript file.</span></span> <span data-ttu-id="9a8f1-132">Необходимо предоставить ссылки на оба этих файла в манифесте, так как платформа Outlook в конечном итоге определяет, следует ли использовать HTML или JavaScript на основе клиента Outlook.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-132">You must provide references to both these files in the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="9a8f1-133">Таким образом, чтобы настроить обработку событий, указите расположение HTML-кода в элементе, а затем в его потомке указите расположение файла JavaScript, на который ссылается `Runtime` `Override` HTML-код.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-133">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="9a8f1-134">Дополнительные информацию о манифестах надстройки Outlook см. в манифестах [надстройки Outlook.](manifests.md)</span><span class="sxs-lookup"><span data-stu-id="9a8f1-134">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="9a8f1-135">Реализация обработки событий</span><span class="sxs-lookup"><span data-stu-id="9a8f1-135">Implement event handling</span></span>

<span data-ttu-id="9a8f1-136">Необходимо реализовать обработку выбранных событий.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-136">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="9a8f1-137">В этом сценарии вы добавим обработку для составления новых элементов.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-137">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="9a8f1-138">В том же проекте быстрого запуска откройте файл **./src/commands/commands.js** в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-138">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="9a8f1-139">После `action` функции вставьте следующие функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-139">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="9a8f1-140">В конце файла добавьте следующие утверждения.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-140">At the end of the file, add the following statements.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

## <a name="try-it-out"></a><span data-ttu-id="9a8f1-141">Проверка</span><span class="sxs-lookup"><span data-stu-id="9a8f1-141">Try it out</span></span>

1. <span data-ttu-id="9a8f1-142">Выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-142">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="9a8f1-143">После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен).</span><span class="sxs-lookup"><span data-stu-id="9a8f1-143">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="9a8f1-144">Чтобы загрузить неопубликованную надстройку в Outlook, следуйте инструкциями из статьи [Загрузка неопубликованных надстроек Outlook для тестирования](sideload-outlook-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="9a8f1-144">Follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md) to sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="9a8f1-145">Создайте новое сообщение в веб-версии Outlook.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-145">In Outlook on the web, create a new message.</span></span>

    ![Снимок экрана: окно сообщения в Outlook в Интернете с темой, настроенной для составить.](../images/outlook-web-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="9a8f1-147">Поведение и ограничения активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="9a8f1-147">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="9a8f1-148">Надстройки, которые активируются на основе событий, рассчитаны на кратковременную работу до 330 секунд.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-148">Add-ins that activate based on events are designed to be short-running, up to 330 seconds only.</span></span> <span data-ttu-id="9a8f1-149">Мы рекомендуем вызвать метод, чтобы надстройка сигнализирует о том, что обработка события `event.completed` запуска завершена.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-149">We recommend you have your add-in call the `event.completed` method to signal it has completed processing the launch event.</span></span> <span data-ttu-id="9a8f1-150">Надстройка также заканчивается, когда пользователь закрывает окно составить.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-150">The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="9a8f1-151">Если у пользователя есть несколько надстройок, которые подписаны на одно и то же событие, платформа Outlook запускает надстройки в определенном порядке.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-151">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="9a8f1-152">В настоящее время можно активно запускать только пять надстройок на основе событий.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-152">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="9a8f1-153">Все дополнительные надстройки добавляются в очередь, а затем запускаются по мере завершения или отключения ранее активных надстроек.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-153">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="9a8f1-154">Пользователь может переключиться или перейти от текущего почтового элемента, в котором запущена надстройка.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-154">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="9a8f1-155">Запущенная надстройка завершит свою работу в фоновом режиме.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-155">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="9a8f1-156">Некоторые Office.js API, которые изменяют или изменяют пользовательский интерфейс, не разрешены из надстройки на основе событий. Ниже следующую ссылку: заблокированные API.</span><span class="sxs-lookup"><span data-stu-id="9a8f1-156">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs.</span></span>

- <span data-ttu-id="9a8f1-157">Under `Office.context.mailbox` :</span><span class="sxs-lookup"><span data-stu-id="9a8f1-157">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="9a8f1-158">Under `Office.context.ui` :</span><span class="sxs-lookup"><span data-stu-id="9a8f1-158">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`
- <span data-ttu-id="9a8f1-159">Under `Office.context.auth` :</span><span class="sxs-lookup"><span data-stu-id="9a8f1-159">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`

## <a name="see-also"></a><span data-ttu-id="9a8f1-160">См. также</span><span class="sxs-lookup"><span data-stu-id="9a8f1-160">See also</span></span>

[<span data-ttu-id="9a8f1-161">Манифесты надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="9a8f1-161">Outlook add-in manifests</span></span>](manifests.md)
