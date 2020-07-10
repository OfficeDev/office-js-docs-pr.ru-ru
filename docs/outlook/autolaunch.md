---
title: Настройка надстройки Outlook для активации на основе событий (Предварительная версия)
description: Узнайте, как настроить надстройку Outlook для активации на основе событий.
ms.topic: article
ms.date: 05/22/2020
localization_priority: Normal
ms.openlocfilehash: 7b14dfb425223dbf486173e7df103d1a2cd09641
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093429"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="63328-103">Настройка надстройки Outlook для активации на основе событий (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="63328-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="63328-104">Без функции активации на основе событий пользователю необходимо явным образом запустить надстройку для выполнения своих задач.</span><span class="sxs-lookup"><span data-stu-id="63328-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="63328-105">Эта функция позволяет надстройке запускать задачи на основе определенных событий, особенно для операций, которые применяются к каждому элементу.</span><span class="sxs-lookup"><span data-stu-id="63328-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="63328-106">Также можно выполнить интеграцию с областью задач и функциональностью без пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="63328-106">You can also integrate with the task pane and UI-less functionality.</span></span> <span data-ttu-id="63328-107">В настоящее время поддерживаются следующие события.</span><span class="sxs-lookup"><span data-stu-id="63328-107">At present, the supported events are as follows.</span></span>

- <span data-ttu-id="63328-108">`OnNewMessageCompose`: На составление нового сообщения (включая ответ, ответить всем и пересылать)</span><span class="sxs-lookup"><span data-stu-id="63328-108">`OnNewMessageCompose`: On composing a new message (includes reply, reply all, and forward)</span></span>
- <span data-ttu-id="63328-109">`OnNewAppointmentOrganizer`: При создании новой встречи</span><span class="sxs-lookup"><span data-stu-id="63328-109">`OnNewAppointmentOrganizer`: On creating a new appointment</span></span>

<span data-ttu-id="63328-110">По завершении этого пошагового руководства у вас будет надстройка, которая запускается при создании нового сообщения.</span><span class="sxs-lookup"><span data-stu-id="63328-110">By the end of this walkthrough, you'll have an add-in that runs whenever a new message is created.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="63328-111">Эта функция поддерживается только для [предварительного просмотра](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) в Outlook в Интернете с подпиской на Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="63328-111">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web with a Microsoft 365 subscription.</span></span> <span data-ttu-id="63328-112">Узнайте [, как просмотреть функцию активации на основе событий,](#how-to-preview-the-event-based-activation-feature) приведенную в этой статье, для получения дополнительных сведений.</span><span class="sxs-lookup"><span data-stu-id="63328-112">See [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="63328-113">Так как функции предварительного просмотра могут быть изменены без предварительного уведомления, они не должны использоваться в производственных надстройках.</span><span class="sxs-lookup"><span data-stu-id="63328-113">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="63328-114">Предварительный просмотр функции активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="63328-114">How to preview the event-based activation feature</span></span>

<span data-ttu-id="63328-115">Мы приглашаем вас испытать функцию активации на основе событий!</span><span class="sxs-lookup"><span data-stu-id="63328-115">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="63328-116">Сообщите нам о своих сценариях и способах их усовершенствования, предоставив отзыв на сайте GitHub (обратитесь к разделу **Отзывы** в конце этой страницы).</span><span class="sxs-lookup"><span data-stu-id="63328-116">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="63328-117">Чтобы просмотреть эту функцию, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="63328-117">To preview this feature:</span></span>

- <span data-ttu-id="63328-118">Ссылка на **бета-** библиотеку в сети CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="63328-118">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="63328-119">[Файл определения типа](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) для компиляции TypeScript и IntelliSense находится в сети CDN и [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="63328-119">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="63328-120">Вы можете установить эти типы с помощью `npm install --save-dev @types/office-js-preview` .</span><span class="sxs-lookup"><span data-stu-id="63328-120">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="63328-121">Запросите доступ к предварительной версии BITS для Outlook в Интернете, используя свою учетную запись Microsoft 365, заполнив и отправив [эту форму запроса](https://aka.ms/OWAPreview).</span><span class="sxs-lookup"><span data-stu-id="63328-121">Request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this request form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="63328-122">Мы будем знать, когда ваш клиент готов.</span><span class="sxs-lookup"><span data-stu-id="63328-122">We'll let you know when your tenant is ready.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="63328-123">Настройка среды</span><span class="sxs-lookup"><span data-stu-id="63328-123">Set up your environment</span></span>

<span data-ttu-id="63328-124">Завершите работу с [быстрым запуском Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) , который создает проект надстройки с помощью генератора Yeoman для надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="63328-124">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="63328-125">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="63328-125">Configure the manifest</span></span>

<span data-ttu-id="63328-126">Чтобы включить активацию надстройки на основе событий, необходимо настроить элемент [среды выполнения](../reference/manifest/runtimes.md) и точку расширения [лаунчевент](../reference/manifest/extensionpoint.md#launchevent-preview) в манифесте.</span><span class="sxs-lookup"><span data-stu-id="63328-126">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the manifest.</span></span> <span data-ttu-id="63328-127">Пока `DesktopFormFactor` это единственный поддерживаемый конструктивный параметр.</span><span class="sxs-lookup"><span data-stu-id="63328-127">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="63328-128">В редакторе кода откройте Быстрый запуск проекта.</span><span class="sxs-lookup"><span data-stu-id="63328-128">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="63328-129">Откройте файл **manifest.xml** , расположенный в корневом каталоге проекта.</span><span class="sxs-lookup"><span data-stu-id="63328-129">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="63328-130">Выберите весь `<VersionOverrides>` узел (включая открывающие и закрывающие теги) и замените его следующим XML-документом.</span><span class="sxs-lookup"><span data-stu-id="63328-130">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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

<span data-ttu-id="63328-131">Outlook в Windows использует файл JavaScript, в то время как Outlook в Интернете использует HTML-файл, который ссылается на тот же файл JavaScript.</span><span class="sxs-lookup"><span data-stu-id="63328-131">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that references the same JavaScript file.</span></span> <span data-ttu-id="63328-132">Необходимо предоставить ссылки на эти файлы в манифесте, так как платформа Outlook в конечном итоге определяет, следует ли использовать HTML или JavaScript на основе клиента Outlook.</span><span class="sxs-lookup"><span data-stu-id="63328-132">You must provide references to both these files in the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="63328-133">Таким образом, чтобы настроить обработку событий, укажите расположение HTML-кода в `Runtime` элементе, а затем в `Override` дочернем элементе укажите расположение файла JavaScript, встроенного или ссылающегося на HTML.</span><span class="sxs-lookup"><span data-stu-id="63328-133">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="63328-134">Чтобы узнать больше о манифестах для надстроек Outlook, ознакомьтесь с разделом [манифесты надстроек Outlook](manifests.md).</span><span class="sxs-lookup"><span data-stu-id="63328-134">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="63328-135">Реализация обработки событий</span><span class="sxs-lookup"><span data-stu-id="63328-135">Implement event handling</span></span>

<span data-ttu-id="63328-136">Необходимо реализовать обработку выбранных событий.</span><span class="sxs-lookup"><span data-stu-id="63328-136">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="63328-137">В этом сценарии вы добавите обработку для создания новых элементов.</span><span class="sxs-lookup"><span data-stu-id="63328-137">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="63328-138">В проекте быстрого запуска откройте **commands.jsфайл./СРК/коммандс/** в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="63328-138">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="63328-139">После `action` функции вставьте следующие функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="63328-139">After the `action` function, insert the following JavaScript functions.</span></span>

    ```js
    function onMessageComposeHandler(event) {
      setSubject();
      event.completed();
    }
    function onAppointmentComposeHandler(event) {
      setSubject();
      event.completed();
    }
    function setSubject() {
      Office.context.mailbox.item.subject.setAsync("Set by an event-based add-in!");
    }
    ```

1. <span data-ttu-id="63328-140">В конце файла добавьте указанные ниже операторы.</span><span class="sxs-lookup"><span data-stu-id="63328-140">At the end of the file, add the following statements.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

## <a name="try-it-out"></a><span data-ttu-id="63328-141">Проверка</span><span class="sxs-lookup"><span data-stu-id="63328-141">Try it out</span></span>

1. <span data-ttu-id="63328-142">Выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="63328-142">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="63328-143">После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен).</span><span class="sxs-lookup"><span data-stu-id="63328-143">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="63328-144">Чтобы загрузить неопубликованную надстройку в Outlook, следуйте инструкциями из статьи [Загрузка неопубликованных надстроек Outlook для тестирования](sideload-outlook-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="63328-144">Follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md) to sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="63328-145">Создайте новое сообщение в веб-версии Outlook.</span><span class="sxs-lookup"><span data-stu-id="63328-145">In Outlook on the web, create a new message.</span></span>

    ![Снимок экрана с окном сообщения в Outlook в Интернете с набором тем для создания.](../images/outlook-web-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="63328-147">Поведение и ограничения активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="63328-147">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="63328-148">Надстройки, активируемые на основе событий, ориентированы на короткий запуск и только до 330 секунд.</span><span class="sxs-lookup"><span data-stu-id="63328-148">Add-ins that activate based on events are designed to be short-running, up to 330 seconds only.</span></span> <span data-ttu-id="63328-149">Мы рекомендуем, чтобы ваша надстройка вызывала `event.completed` метод, чтобы сообщить, что обработка события запуска завершена.</span><span class="sxs-lookup"><span data-stu-id="63328-149">We recommend you have your add-in call the `event.completed` method to signal it has completed processing the launch event.</span></span> <span data-ttu-id="63328-150">Кроме того, надстройка завершает работу, когда пользователь закрывает окно создания.</span><span class="sxs-lookup"><span data-stu-id="63328-150">The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="63328-151">Если у пользователя есть несколько надстроек, подписанных на одно и то же событие, платформа Outlook запускает надстройку в неопределенном порядке.</span><span class="sxs-lookup"><span data-stu-id="63328-151">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="63328-152">В настоящее время только пять надстроек на основе событий могут быть запущены в активном состоянии.</span><span class="sxs-lookup"><span data-stu-id="63328-152">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="63328-153">Все дополнительные надстройки помещаются в очередь, а затем выполняются, как только ранее активные надстройки завершаются или отключаются.</span><span class="sxs-lookup"><span data-stu-id="63328-153">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="63328-154">Пользователь может переключить или покинуть текущий почтовый элемент, где запущена надстройка.</span><span class="sxs-lookup"><span data-stu-id="63328-154">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="63328-155">Запущенная надстройка завершит свою работу в фоновом режиме.</span><span class="sxs-lookup"><span data-stu-id="63328-155">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="63328-156">Некоторые API Office.js, которые изменяют или изменяют пользовательский интерфейс, не поддерживаются в надстройках, основанных на событиях. Ниже приведены заблокированные API.</span><span class="sxs-lookup"><span data-stu-id="63328-156">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs.</span></span>

- <span data-ttu-id="63328-157">В разделе `Office.context.mailbox` :</span><span class="sxs-lookup"><span data-stu-id="63328-157">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="63328-158">В разделе `Office.context.ui` :</span><span class="sxs-lookup"><span data-stu-id="63328-158">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`
- <span data-ttu-id="63328-159">В разделе `Office.context.auth` :</span><span class="sxs-lookup"><span data-stu-id="63328-159">Under `Office.context.auth`:</span></span>
  - `getAccessTokenAsync`

## <a name="see-also"></a><span data-ttu-id="63328-160">См. также</span><span class="sxs-lookup"><span data-stu-id="63328-160">See also</span></span>

[<span data-ttu-id="63328-161">Манифесты надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="63328-161">Outlook add-in manifests</span></span>](manifests.md)
