---
title: Настройка надстройки Outlook для активации на основе событий (предварительный просмотр)
description: Узнайте, как настроить Outlook надстройку для активации на основе событий.
ms.topic: article
ms.date: 05/04/2021
localization_priority: Normal
ms.openlocfilehash: 0052f08e9c6a3903f4adb48efca3ff29a6d21467
ms.sourcegitcommit: 8fbc7c7eb47875bf022e402b13858695a8536ec5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2021
ms.locfileid: "52253326"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="317a9-103">Настройка надстройки Outlook для активации на основе событий (предварительный просмотр)</span><span class="sxs-lookup"><span data-stu-id="317a9-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="317a9-104">Без функции активации на основе событий пользователю необходимо явно запустить надстройки для выполнения задач.</span><span class="sxs-lookup"><span data-stu-id="317a9-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="317a9-105">Эта функция позволяет надстройки выполнять задачи на основе определенных событий, особенно для операций, применимых к каждому элементу.</span><span class="sxs-lookup"><span data-stu-id="317a9-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="317a9-106">Вы также можете интегрироваться с области задач и функциональными возможностями без пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="317a9-106">You can also integrate with the task pane and UI-less functionality.</span></span>

<span data-ttu-id="317a9-107">К концу этого погона у вас будет надстройка, которая запускается всякий раз, когда создается новый элемент и задает объект.</span><span class="sxs-lookup"><span data-stu-id="317a9-107">By the end of this walkthrough, you'll have an add-in that runs whenever a new item is created and sets the subject.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="317a9-108">Эта функция поддерживается [](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) только для предварительного просмотра Outlook в Интернете и Windows с Microsoft 365 подпиской.</span><span class="sxs-lookup"><span data-stu-id="317a9-108">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="317a9-109">Дополнительные сведения см. [в статье Как просмотреть](#how-to-preview-the-event-based-activation-feature) функцию активации на основе событий в этой статье.</span><span class="sxs-lookup"><span data-stu-id="317a9-109">For more details, see [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article.</span></span>
>
> <span data-ttu-id="317a9-110">Так как функции предварительного просмотра могут изменяться без предварительного уведомления, их не следует использовать в надстройки производства.</span><span class="sxs-lookup"><span data-stu-id="317a9-110">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="supported-events"></a><span data-ttu-id="317a9-111">Поддерживаемые события</span><span class="sxs-lookup"><span data-stu-id="317a9-111">Supported events</span></span>

<span data-ttu-id="317a9-112">В настоящее время поддерживаются следующие события.</span><span class="sxs-lookup"><span data-stu-id="317a9-112">At present, the following events are supported.</span></span>

|<span data-ttu-id="317a9-113">Событие</span><span class="sxs-lookup"><span data-stu-id="317a9-113">Event</span></span>|<span data-ttu-id="317a9-114">Описание</span><span class="sxs-lookup"><span data-stu-id="317a9-114">Description</span></span>|<span data-ttu-id="317a9-115">Клиенты</span><span class="sxs-lookup"><span data-stu-id="317a9-115">Clients</span></span>|
|---|---|---|
|`OnNewMessageCompose`|<span data-ttu-id="317a9-116">При составлении нового сообщения (включает ответ, ответ все и вперед), но не при редактировании, например, черновика.</span><span class="sxs-lookup"><span data-stu-id="317a9-116">On composing a new message (includes reply, reply all, and forward) but not on editing, for example, a draft.</span></span>|<span data-ttu-id="317a9-117">Windows, веб</span><span class="sxs-lookup"><span data-stu-id="317a9-117">Windows, web</span></span>|
|`OnNewAppointmentOrganizer`|<span data-ttu-id="317a9-118">О создании новой встречи, но не о редактировании существующего.</span><span class="sxs-lookup"><span data-stu-id="317a9-118">On creating a new appointment but not on editing an existing one.</span></span>|<span data-ttu-id="317a9-119">Windows, веб</span><span class="sxs-lookup"><span data-stu-id="317a9-119">Windows, web</span></span>|
|`OnMessageAttachmentsChanged`|<span data-ttu-id="317a9-120">При добавлении или удалении вложений при сочинении сообщения.</span><span class="sxs-lookup"><span data-stu-id="317a9-120">On adding or removing attachments while composing a message.</span></span>|<span data-ttu-id="317a9-121">Windows</span><span class="sxs-lookup"><span data-stu-id="317a9-121">Windows</span></span>|
|`OnAppointmentAttachmentsChanged`|<span data-ttu-id="317a9-122">При добавлении или удалении вложений при записи на прием.</span><span class="sxs-lookup"><span data-stu-id="317a9-122">On adding or removing attachments while composing an appointment.</span></span>|<span data-ttu-id="317a9-123">Windows</span><span class="sxs-lookup"><span data-stu-id="317a9-123">Windows</span></span>|
|`OnMessageRecipientsChanged`|<span data-ttu-id="317a9-124">При добавлении или удалении получателей при сочинении сообщения.</span><span class="sxs-lookup"><span data-stu-id="317a9-124">On adding or removing recipients while composing a message.</span></span>|<span data-ttu-id="317a9-125">Windows</span><span class="sxs-lookup"><span data-stu-id="317a9-125">Windows</span></span>|
|`OnAppointmentAttendeesChanged`|<span data-ttu-id="317a9-126">При добавлении или удалении участников при записи на прием.</span><span class="sxs-lookup"><span data-stu-id="317a9-126">On adding or removing attendees while composing an appointment.</span></span>|<span data-ttu-id="317a9-127">Windows</span><span class="sxs-lookup"><span data-stu-id="317a9-127">Windows</span></span>|
|`OnAppointmentTimeChanged`|<span data-ttu-id="317a9-128">При изменении даты и времени при записи на прием.</span><span class="sxs-lookup"><span data-stu-id="317a9-128">On changing date/time while composing an appointment.</span></span>|<span data-ttu-id="317a9-129">Windows</span><span class="sxs-lookup"><span data-stu-id="317a9-129">Windows</span></span>|
|`OnAppointmentRecurrenceChanged`|<span data-ttu-id="317a9-130">При добавлении, изменении или удалении сведений о повторении при записи на прием.</span><span class="sxs-lookup"><span data-stu-id="317a9-130">On adding, changing, or removing the recurrence details while composing an appointment.</span></span> <span data-ttu-id="317a9-131">Если дата и время изменены, `OnAppointmentTimeChanged` событие также будет уволено.</span><span class="sxs-lookup"><span data-stu-id="317a9-131">If the date/time is changed, the `OnAppointmentTimeChanged` event will also be fired.</span></span>|<span data-ttu-id="317a9-132">Windows</span><span class="sxs-lookup"><span data-stu-id="317a9-132">Windows</span></span>|
|`OnInfoBarDismissClicked`|<span data-ttu-id="317a9-133">При отклонении уведомления при записи сообщения или элемента встречи.</span><span class="sxs-lookup"><span data-stu-id="317a9-133">On dismissing a notification while composing a message or appointment item.</span></span> <span data-ttu-id="317a9-134">Уведомления будут получать только надстройка, которая добавила уведомление.</span><span class="sxs-lookup"><span data-stu-id="317a9-134">Only the add-in that added the notification will be notified.</span></span>|<span data-ttu-id="317a9-135">Windows</span><span class="sxs-lookup"><span data-stu-id="317a9-135">Windows</span></span>|

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="317a9-136">Просмотр функции активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="317a9-136">How to preview the event-based activation feature</span></span>

<span data-ttu-id="317a9-137">Мы приглашаем вас попробовать функцию активации на основе событий!</span><span class="sxs-lookup"><span data-stu-id="317a9-137">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="317a9-138">Дайте нам знать о ваших сценариях и о том, как мы можем улучшить ситуацию, GitHub с помощью GitHub (см. раздел **Обратная** связь в конце этой страницы).</span><span class="sxs-lookup"><span data-stu-id="317a9-138">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="317a9-139">Чтобы просмотреть эту функцию:</span><span class="sxs-lookup"><span data-stu-id="317a9-139">To preview this feature:</span></span>

- <span data-ttu-id="317a9-140">Для Outlook в Интернете:</span><span class="sxs-lookup"><span data-stu-id="317a9-140">For Outlook on the web:</span></span>
  - <span data-ttu-id="317a9-141">[Настройка целевого выпуска для](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)Microsoft 365 клиента.</span><span class="sxs-lookup"><span data-stu-id="317a9-141">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="317a9-142">Ссылка  на бета-библиотеку на CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="317a9-142">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="317a9-143">Файл [определения типа](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) для компиляции и IntelliSense typeScript CDN и [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="317a9-143">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="317a9-144">Эти типы можно установить с `npm install --save-dev @types/office-js-preview` помощью .</span><span class="sxs-lookup"><span data-stu-id="317a9-144">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="317a9-145">Для Outlook на Windows:</span><span class="sxs-lookup"><span data-stu-id="317a9-145">For Outlook on Windows:</span></span>
  - <span data-ttu-id="317a9-146">Минимальная требуемая сборка — 16.0.14026.20000.</span><span class="sxs-lookup"><span data-stu-id="317a9-146">The minimum required build is 16.0.14026.20000.</span></span> <span data-ttu-id="317a9-147">Присоединяйтесь [к Office программы insider](https://insider.office.com) для доступа к Office бета-сборки.</span><span class="sxs-lookup"><span data-stu-id="317a9-147">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>
  - <span data-ttu-id="317a9-148">Настройка реестра:</span><span class="sxs-lookup"><span data-stu-id="317a9-148">Configure the registry:</span></span>
    1. <span data-ttu-id="317a9-149">Создание ключа `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer` реестра.</span><span class="sxs-lookup"><span data-stu-id="317a9-149">Create the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer`.</span></span>
    1. <span data-ttu-id="317a9-150">Добавьте запись с `EnableBetaAPIsInJavaScript` именем и установите значение `1` .</span><span class="sxs-lookup"><span data-stu-id="317a9-150">Add an entry named `EnableBetaAPIsInJavaScript` and set the value to `1`.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="317a9-151">Настройка среды</span><span class="sxs-lookup"><span data-stu-id="317a9-151">Set up your environment</span></span>

<span data-ttu-id="317a9-152">Выполните [Outlook,](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) который создает проект надстройки с генератором Yeoman для Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="317a9-152">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="317a9-153">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="317a9-153">Configure the manifest</span></span>

<span data-ttu-id="317a9-154">Чтобы включить активацию надстройки на основе событий, необходимо настроить элемент [Runtimes](../reference/manifest/runtimes.md) и точку расширения [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) в узле `VersionOverridesV1_1` манифеста.</span><span class="sxs-lookup"><span data-stu-id="317a9-154">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="317a9-155">Пока это `DesktopFormFactor` единственный поддерживаемый форм-фактор.</span><span class="sxs-lookup"><span data-stu-id="317a9-155">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="317a9-156">В редакторе кода откройте проект быстрого запуска.</span><span class="sxs-lookup"><span data-stu-id="317a9-156">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="317a9-157">Откройте **manifest.xml** файл, расположенный в корне проекта.</span><span class="sxs-lookup"><span data-stu-id="317a9-157">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="317a9-158">Выберите весь узел (включая открытые и закрываемые теги) и замените его на следующий XML, а затем `<VersionOverrides>` сохраните изменения.</span><span class="sxs-lookup"><span data-stu-id="317a9-158">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.</span></span>

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

<span data-ttu-id="317a9-159">Outlook на Windows использует файл JavaScript, а Outlook в Интернете использует HTML-файл, который может ссылаться на один и тот же файл JavaScript.</span><span class="sxs-lookup"><span data-stu-id="317a9-159">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="317a9-160">Необходимо предоставить ссылки на оба этих файла в узле манифеста, так как платформа Outlook в конечном счете определяет, следует ли использовать HTML или JavaScript на основе Outlook `Resources` клиента.</span><span class="sxs-lookup"><span data-stu-id="317a9-160">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="317a9-161">Таким образом, чтобы настроить обработку событий, укадь расположение HTML в элементе, а затем в его детском элементе укаймляй расположение файла JavaScript, вписаного или ссылаемого `Runtime` `Override` HTML.</span><span class="sxs-lookup"><span data-stu-id="317a9-161">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="317a9-162">Дополнительные информацию о манифестах для Outlook надстройки см. в Outlook [манифестах надстройки.](manifests.md)</span><span class="sxs-lookup"><span data-stu-id="317a9-162">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="317a9-163">Реализация обработки событий</span><span class="sxs-lookup"><span data-stu-id="317a9-163">Implement event handling</span></span>

<span data-ttu-id="317a9-164">Для выбранных событий необходимо реализовать обработку.</span><span class="sxs-lookup"><span data-stu-id="317a9-164">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="317a9-165">В этом сценарии вы добавим обработку для составления новых элементов.</span><span class="sxs-lookup"><span data-stu-id="317a9-165">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="317a9-166">В том же проекте быстрого запуска откройте **файл ./src/commands/commands.js** в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="317a9-166">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="317a9-167">После `action` функции вставьте следующие функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="317a9-167">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="317a9-168">Чтобы функции работали  Outlook в Интернете с этим проектом, созданным генератором Yeoman для Office надстройки, добавьте следующие утверждения в конце файла.</span><span class="sxs-lookup"><span data-stu-id="317a9-168">For the functions to work in **Outlook on the web** with this project generated by the Yeoman generator for Office Add-ins, add the following statements at the end of the file.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

1. <span data-ttu-id="317a9-169">Чтобы функции работали **в Outlook Windows, добавьте** следующий код JavaScript в конце файла.</span><span class="sxs-lookup"><span data-stu-id="317a9-169">For the functions to work in **Outlook on Windows**, add the following JavaScript code at the end of the file.</span></span>

    ```js
    if (Office.actions) {
      // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
      Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
      Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    }
    ```

    <span data-ttu-id="317a9-170">**Примечание.** Проверка `Office.actions` на то, что Outlook в Интернете игнорирует эти утверждения.</span><span class="sxs-lookup"><span data-stu-id="317a9-170">**Note**: Checking for `Office.actions` ensures that Outlook on the web ignores these statements.</span></span>

1. <span data-ttu-id="317a9-171">Сохраните изменения.</span><span class="sxs-lookup"><span data-stu-id="317a9-171">Save your changes.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="317a9-172">Проверка</span><span class="sxs-lookup"><span data-stu-id="317a9-172">Try it out</span></span>

1. <span data-ttu-id="317a9-173">Выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="317a9-173">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="317a9-174">После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен) и будет загружена ваша неопубликованная надстройка.</span><span class="sxs-lookup"><span data-stu-id="317a9-174">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

1. <span data-ttu-id="317a9-175">Создайте новое сообщение в веб-версии Outlook.</span><span class="sxs-lookup"><span data-stu-id="317a9-175">In Outlook on the web, create a new message.</span></span>

    ![Снимок экрана окна сообщения в Outlook веб-страницы с набором субъекта на композит](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="317a9-177">В Outlook Windows создайте новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="317a9-177">In Outlook on Windows, create a new message.</span></span>

    ![Снимок экрана окна сообщения в Outlook Windows с набором субъекта на композицию](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > <span data-ttu-id="317a9-179">Если вы видите ошибку "Мы не можем открыть эту надстройку из localhost", необходимо включить освобождение от циклов.</span><span class="sxs-lookup"><span data-stu-id="317a9-179">If you see the error "We can't open this add-in from localhost," you'll need to enable a loopback exemption.</span></span>
    >
    > 1. <span data-ttu-id="317a9-180">Закройте Outlook.</span><span class="sxs-lookup"><span data-stu-id="317a9-180">Close Outlook.</span></span>
    > 2. <span data-ttu-id="317a9-181">Откройте диспетчер **задач** и убедитесь, что **msoadfs.exe** процесс не запущен.</span><span class="sxs-lookup"><span data-stu-id="317a9-181">Open the **Task Manager** and ensure that the **msoadfs.exe** process is not running.</span></span>
    > 3. <span data-ttu-id="317a9-182">Выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="317a9-182">Run the following command.</span></span>
    >
    >     ```command&nbsp;line
    >     call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >     ```
    >
    > 4. <span data-ttu-id="317a9-183">Перезапустите Outlook.</span><span class="sxs-lookup"><span data-stu-id="317a9-183">Restart Outlook.</span></span>

## <a name="debug"></a><span data-ttu-id="317a9-184">Debug</span><span class="sxs-lookup"><span data-stu-id="317a9-184">Debug</span></span>

<span data-ttu-id="317a9-185">При реализации собственных функций может потребоваться отламывка кода.</span><span class="sxs-lookup"><span data-stu-id="317a9-185">As you implement your own functionality, you may need to debug your code.</span></span> <span data-ttu-id="317a9-186">Инструкции по отламывить активацию надстройки на основе событий см. в Outlook [событиях.](debug-autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="317a9-186">For guidance on how to debug event-based add-in activation, see [Debug your event-based Outlook add-in](debug-autolaunch.md).</span></span>

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="317a9-187">Поведение и ограничения активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="317a9-187">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="317a9-188">Надстройки, которые активируются на основе событий, как ожидается, будут короткими, легкими и максимально неинвазивными.</span><span class="sxs-lookup"><span data-stu-id="317a9-188">Add-ins that activate based on events are expected to be short-running, lightweight, and as non-invasive as possible.</span></span> <span data-ttu-id="317a9-189">Чтобы сигнализировать, что надстройка завершила обработку события запуска, рекомендуется использовать метод вызова `event.completed` надстройки.</span><span class="sxs-lookup"><span data-stu-id="317a9-189">To signal that your add-in has completed processing the launch event, we recommend you have your add-in call the `event.completed` method.</span></span> <span data-ttu-id="317a9-190">Если этот вызов не будет выполнен, надстройка будет работать в течение примерно 300 секунд, что является максимальным сроком, разрешенным для запуска надстроек на основе событий. Надстройка также заканчивается, когда пользователь закрывает окно записи.</span><span class="sxs-lookup"><span data-stu-id="317a9-190">If that call is not made, the add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="317a9-191">Если у пользователя есть несколько надстройок, которые подписаны на одно и то же событие, Outlook платформа запускает надстройки без определенного порядка.</span><span class="sxs-lookup"><span data-stu-id="317a9-191">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="317a9-192">В настоящее время можно активно запускать только пять надстройок на основе событий.</span><span class="sxs-lookup"><span data-stu-id="317a9-192">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="317a9-193">Все дополнительные надстройки отодвигаются в очередь, а затем запускаются по мере завершения или отключения ранее активных надстроек.</span><span class="sxs-lookup"><span data-stu-id="317a9-193">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="317a9-194">Пользователь может переключаться или перемещаться от текущего элемента почты, где надстройка начала работать.</span><span class="sxs-lookup"><span data-stu-id="317a9-194">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="317a9-195">Запущенная надстройка завершит свою работу в фоновом режиме.</span><span class="sxs-lookup"><span data-stu-id="317a9-195">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="317a9-196">Некоторые Office.js API, которые изменяют или изменяют пользовательский интерфейс, не допускаются из надстройок на основе событий. Следующие API заблокированы:</span><span class="sxs-lookup"><span data-stu-id="317a9-196">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="317a9-197">В `Office.context.auth` статье:</span><span class="sxs-lookup"><span data-stu-id="317a9-197">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`
- <span data-ttu-id="317a9-198">В `Office.context.mailbox` статье:</span><span class="sxs-lookup"><span data-stu-id="317a9-198">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="317a9-199">В `Office.context.mailbox.item` статье:</span><span class="sxs-lookup"><span data-stu-id="317a9-199">Under `Office.context.mailbox.item`:</span></span>
  - `close`
- <span data-ttu-id="317a9-200">В `Office.context.ui` статье:</span><span class="sxs-lookup"><span data-stu-id="317a9-200">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a><span data-ttu-id="317a9-201">См. также</span><span class="sxs-lookup"><span data-stu-id="317a9-201">See also</span></span>

- [<span data-ttu-id="317a9-202">Манифесты надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="317a9-202">Outlook add-in manifests</span></span>](manifests.md)
- [<span data-ttu-id="317a9-203">Отламывка надстроек на основе событий</span><span class="sxs-lookup"><span data-stu-id="317a9-203">How to debug event-based add-ins</span></span>](debug-autolaunch.md)
