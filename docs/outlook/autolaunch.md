---
title: Настройте Outlook для активации на основе событий (предварительный просмотр)
description: Узнайте, как настроить Outlook для активации на основе событий.
ms.topic: article
ms.date: 05/18/2021
localization_priority: Normal
ms.openlocfilehash: 721f05e1c835e066744598ecb2bd416c6a6b0526
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555242"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="47725-103">Настройте Outlook для активации на основе событий (предварительный просмотр)</span><span class="sxs-lookup"><span data-stu-id="47725-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="47725-104">Без функции активации на основе событий пользователь должен явно запустить надстройу для выполнения своих задач.</span><span class="sxs-lookup"><span data-stu-id="47725-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="47725-105">Эта функция позволяет надстройки выполнять задачи на основе определенных событий, особенно для операций, которые применяются к каждому элементу.</span><span class="sxs-lookup"><span data-stu-id="47725-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="47725-106">Вы также можете интегрироваться с функцией панели задач и пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="47725-106">You can also integrate with the task pane and UI-less functionality.</span></span>

<span data-ttu-id="47725-107">К концу этого пошагового руководства, вы будете иметь надстройку, которая работает всякий раз, когда новый элемент создается и устанавливает тему.</span><span class="sxs-lookup"><span data-stu-id="47725-107">By the end of this walkthrough, you'll have an add-in that runs whenever a new item is created and sets the subject.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="47725-108">Эта функция поддерживается только для [предварительного](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) просмотра Outlook веб-сайтах и Windows с Microsoft 365 подпиской.</span><span class="sxs-lookup"><span data-stu-id="47725-108">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="47725-109">Для получения более подробной информации [в этой статье узнайте, как просмотреть функцию активации на](#how-to-preview-the-event-based-activation-feature) основе событий.</span><span class="sxs-lookup"><span data-stu-id="47725-109">For more details, see [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article.</span></span>
>
> <span data-ttu-id="47725-110">Поскольку функции предварительного просмотра могут быть изменения без предварительного уведомления, они не должны использоваться в производственных дополнениях.</span><span class="sxs-lookup"><span data-stu-id="47725-110">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="supported-events"></a><span data-ttu-id="47725-111">Поддерживаемые события</span><span class="sxs-lookup"><span data-stu-id="47725-111">Supported events</span></span>

<span data-ttu-id="47725-112">В настоящее время поддерживаются следующие мероприятия.</span><span class="sxs-lookup"><span data-stu-id="47725-112">At present, the following events are supported.</span></span>

|<span data-ttu-id="47725-113">Событие</span><span class="sxs-lookup"><span data-stu-id="47725-113">Event</span></span>|<span data-ttu-id="47725-114">Описание</span><span class="sxs-lookup"><span data-stu-id="47725-114">Description</span></span>|<span data-ttu-id="47725-115">Клиенты</span><span class="sxs-lookup"><span data-stu-id="47725-115">Clients</span></span>|
|---|---|---|
|`OnNewMessageCompose`|<span data-ttu-id="47725-116">О составлении нового сообщения (включает ответ, ответьте все и вперед), но не о редактировании, например, проекта.</span><span class="sxs-lookup"><span data-stu-id="47725-116">On composing a new message (includes reply, reply all, and forward) but not on editing, for example, a draft.</span></span>|<span data-ttu-id="47725-117">Windows, веб</span><span class="sxs-lookup"><span data-stu-id="47725-117">Windows, web</span></span>|
|`OnNewAppointmentOrganizer`|<span data-ttu-id="47725-118">О создании новой встречи, но не о редактировании существующей.</span><span class="sxs-lookup"><span data-stu-id="47725-118">On creating a new appointment but not on editing an existing one.</span></span>|<span data-ttu-id="47725-119">Windows, веб</span><span class="sxs-lookup"><span data-stu-id="47725-119">Windows, web</span></span>|
|`OnMessageAttachmentsChanged`|<span data-ttu-id="47725-120">При добавлении или удалении вложений при составлении сообщения.</span><span class="sxs-lookup"><span data-stu-id="47725-120">On adding or removing attachments while composing a message.</span></span>|<span data-ttu-id="47725-121">Windows</span><span class="sxs-lookup"><span data-stu-id="47725-121">Windows</span></span>|
|`OnAppointmentAttachmentsChanged`|<span data-ttu-id="47725-122">При добавлении или удалении вложений при составлении встречи.</span><span class="sxs-lookup"><span data-stu-id="47725-122">On adding or removing attachments while composing an appointment.</span></span>|<span data-ttu-id="47725-123">Windows</span><span class="sxs-lookup"><span data-stu-id="47725-123">Windows</span></span>|
|`OnMessageRecipientsChanged`|<span data-ttu-id="47725-124">При добавлении или удалении получателей при составлении сообщения.</span><span class="sxs-lookup"><span data-stu-id="47725-124">On adding or removing recipients while composing a message.</span></span>|<span data-ttu-id="47725-125">Windows</span><span class="sxs-lookup"><span data-stu-id="47725-125">Windows</span></span>|
|`OnAppointmentAttendeesChanged`|<span data-ttu-id="47725-126">При добавлении или удалении участников при составлении встречи.</span><span class="sxs-lookup"><span data-stu-id="47725-126">On adding or removing attendees while composing an appointment.</span></span>|<span data-ttu-id="47725-127">Windows</span><span class="sxs-lookup"><span data-stu-id="47725-127">Windows</span></span>|
|`OnAppointmentTimeChanged`|<span data-ttu-id="47725-128">При изменении даты/времени при составлении встречи.</span><span class="sxs-lookup"><span data-stu-id="47725-128">On changing date/time while composing an appointment.</span></span>|<span data-ttu-id="47725-129">Windows</span><span class="sxs-lookup"><span data-stu-id="47725-129">Windows</span></span>|
|`OnAppointmentRecurrenceChanged`|<span data-ttu-id="47725-130">При добавлении, изменении или удалении деталей повторения при составлении записи на прием.</span><span class="sxs-lookup"><span data-stu-id="47725-130">On adding, changing, or removing the recurrence details while composing an appointment.</span></span> <span data-ttu-id="47725-131">Если дата/время изменены, `OnAppointmentTimeChanged` событие также будет уволено.</span><span class="sxs-lookup"><span data-stu-id="47725-131">If the date/time is changed, the `OnAppointmentTimeChanged` event will also be fired.</span></span>|<span data-ttu-id="47725-132">Windows</span><span class="sxs-lookup"><span data-stu-id="47725-132">Windows</span></span>|
|`OnInfoBarDismissClicked`|<span data-ttu-id="47725-133">При увольнении уведомления при составлении сообщения или пункта назначения.</span><span class="sxs-lookup"><span data-stu-id="47725-133">On dismissing a notification while composing a message or appointment item.</span></span> <span data-ttu-id="47725-134">Только надстройкое, добавляемое уведомление, будет уведомлено.</span><span class="sxs-lookup"><span data-stu-id="47725-134">Only the add-in that added the notification will be notified.</span></span>|<span data-ttu-id="47725-135">Windows</span><span class="sxs-lookup"><span data-stu-id="47725-135">Windows</span></span>|

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="47725-136">Как просмотреть функцию активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="47725-136">How to preview the event-based activation feature</span></span>

<span data-ttu-id="47725-137">Мы приглашаем Вас опробовать функцию активации на основе событий!</span><span class="sxs-lookup"><span data-stu-id="47725-137">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="47725-138">Сообщите нам о ваших сценариях и о том, как мы можем улучшить их, дав нам обратную связь GitHub **(см.** раздел Обратная связь в конце этой страницы).</span><span class="sxs-lookup"><span data-stu-id="47725-138">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="47725-139">Для просмотра этой функции:</span><span class="sxs-lookup"><span data-stu-id="47725-139">To preview this feature:</span></span>

- <span data-ttu-id="47725-140">Для Outlook в Интернете:</span><span class="sxs-lookup"><span data-stu-id="47725-140">For Outlook on the web:</span></span>
  - <span data-ttu-id="47725-141">[Настройте целевой релиз на Microsoft 365 арендатора.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="47725-141">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="47725-142">Ссылка  на бета-библиотеку на CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="47725-142">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="47725-143">Файл [определения типа для](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) компиляции TypeScript и IntelliSense найден на CDN [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="47725-143">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="47725-144">Вы можете установить эти типы с `npm install --save-dev @types/office-js-preview` .</span><span class="sxs-lookup"><span data-stu-id="47725-144">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="47725-145">Для Outlook на Windows:</span><span class="sxs-lookup"><span data-stu-id="47725-145">For Outlook on Windows:</span></span>
  - <span data-ttu-id="47725-146">Минимальная требуемая сборка составляет 16.0.14026.20000.</span><span class="sxs-lookup"><span data-stu-id="47725-146">The minimum required build is 16.0.14026.20000.</span></span> <span data-ttu-id="47725-147">Присоединяйтесь [к Office Insider для](https://insider.office.com) доступа к бета Office котейной версии.</span><span class="sxs-lookup"><span data-stu-id="47725-147">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>
  - <span data-ttu-id="47725-148">Настройте реестр.</span><span class="sxs-lookup"><span data-stu-id="47725-148">Configure the registry.</span></span> <span data-ttu-id="47725-149">Outlook включает в себя локательную копию производственной и бета-версии Office.js вместо загрузки с CDN.</span><span class="sxs-lookup"><span data-stu-id="47725-149">Outlook includes a local copy of the production and beta versions of Office.js instead of loading from the CDN.</span></span> <span data-ttu-id="47725-150">По умолчанию ссылается локал-производственная копия API.</span><span class="sxs-lookup"><span data-stu-id="47725-150">By default, the local production copy of the API is referenced.</span></span> <span data-ttu-id="47725-151">Чтобы перейти на локаную бета-копию api Outlook JavaScript, необходимо добавить эту запись реестра, в противном случае бета-API могут не быть найдены.</span><span class="sxs-lookup"><span data-stu-id="47725-151">To switch to the local beta copy of the Outlook JavaScript APIs, you need to add this registry entry, otherwise beta APIs may not be found.</span></span>
    1. <span data-ttu-id="47725-152">Создание ключа `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer` реестра.</span><span class="sxs-lookup"><span data-stu-id="47725-152">Create the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer`.</span></span>
    1. <span data-ttu-id="47725-153">Добавьте `EnableBetaAPIsInJavaScript` именуемую запись и установите `1` значение.</span><span class="sxs-lookup"><span data-stu-id="47725-153">Add an entry named `EnableBetaAPIsInJavaScript` and set the value to `1`.</span></span> <span data-ttu-id="47725-154">На приведенном ниже изображении показано, как должен выглядеть реестр.</span><span class="sxs-lookup"><span data-stu-id="47725-154">The following image shows what the registry should look like.</span></span>

        ![Скриншот редактора реестра с ключевым значением реестра EnableBetaAPIsInJavaScript](../images/outlook-beta-registry-key.png)

## <a name="set-up-your-environment"></a><span data-ttu-id="47725-156">Настройка среды</span><span class="sxs-lookup"><span data-stu-id="47725-156">Set up your environment</span></span>

<span data-ttu-id="47725-157">Завершите [Outlook,](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) который создает надстройки проекта с генератором Yeoman для Office дополнительных висел.</span><span class="sxs-lookup"><span data-stu-id="47725-157">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="47725-158">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="47725-158">Configure the manifest</span></span>

<span data-ttu-id="47725-159">Для активации надстройок на основе событий необходимо настроить элемент [Runtimes и](../reference/manifest/runtimes.md) [точку расширения LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) `VersionOverridesV1_1` в узел манифеста.</span><span class="sxs-lookup"><span data-stu-id="47725-159">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="47725-160">На данный `DesktopFormFactor` момент, является единственным поддерживаемым форм-фактором.</span><span class="sxs-lookup"><span data-stu-id="47725-160">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="47725-161">В редакторе кода откройте проект быстрого запуска.</span><span class="sxs-lookup"><span data-stu-id="47725-161">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="47725-162">Откройте **manifest.xml** файл, расположенный в корне вашего проекта.</span><span class="sxs-lookup"><span data-stu-id="47725-162">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="47725-163">Выберите весь `<VersionOverrides>` узел (включая открытые и близкие теги) и замените его следующим XML, а затем сохраните изменения.</span><span class="sxs-lookup"><span data-stu-id="47725-163">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.</span></span>

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
              <!-- Events supported on the web and on Windows. -->
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
              <!-- Events supported only on Windows. -->
              <LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="onMessageAttachmentsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttachmentsChanged" FunctionName="onAppointmentAttachmentsChangedHandler" />
              <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="onMessageRecipientsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttendeesChanged" FunctionName="onAppointmentAttendeesChangedHandler" />
              <LaunchEvent Type="OnAppointmentTimeChanged" FunctionName="onAppointmentTimeChangedHandler" />
              <LaunchEvent Type="OnAppointmentRecurrenceChanged" FunctionName="onAppointmentRecurrenceChangedHandler" />
              <LaunchEvent Type="OnInfoBarDismissClicked" FunctionName="onInfobarDismissClickedHandler" />
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

<span data-ttu-id="47725-164">Outlook на Windows файл JavaScript, в то время как Outlook в Интернете использует HTML файл, который может ссылаться на тот же файл JavaScript.</span><span class="sxs-lookup"><span data-stu-id="47725-164">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="47725-165">Вы должны предоставить ссылки на оба этих файла `Resources` в узел манифеста, как платформа Outlook в конечном итоге определяет, следует ли использовать HTML или JavaScript на основе Outlook клиента.</span><span class="sxs-lookup"><span data-stu-id="47725-165">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="47725-166">Таким образом, чтобы настроить обработку событий, уведите расположение HTML в `Runtime` элементе, затем в `Override` его элементе ребенка предоставьте расположение файла JavaScript, вписанный или на который ссылается HTML.</span><span class="sxs-lookup"><span data-stu-id="47725-166">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="47725-167">Чтобы узнать больше о манифестах Outlook дополнительных надстройок, [Outlook дополнительные дополнения.](manifests.md)</span><span class="sxs-lookup"><span data-stu-id="47725-167">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="47725-168">Реализация обработки событий</span><span class="sxs-lookup"><span data-stu-id="47725-168">Implement event handling</span></span>

<span data-ttu-id="47725-169">Вы должны реализовать обработку выбранных событий.</span><span class="sxs-lookup"><span data-stu-id="47725-169">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="47725-170">В этом сценарии вы добавите обработку для составления новых элементов.</span><span class="sxs-lookup"><span data-stu-id="47725-170">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="47725-171">С того же проекта быстрого запуска откройте файл **./src/commands/commands.js** в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="47725-171">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="47725-172">После `action` функции вставьте следующие функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="47725-172">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="47725-173">Добавьте следующий код JavaScript в конце файла.</span><span class="sxs-lookup"><span data-stu-id="47725-173">Add the following JavaScript code at the end of the file.</span></span>

    ```js
    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
    Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    ```

1. <span data-ttu-id="47725-174">Сохраните изменения.</span><span class="sxs-lookup"><span data-stu-id="47725-174">Save your changes.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="47725-175">Проверка</span><span class="sxs-lookup"><span data-stu-id="47725-175">Try it out</span></span>

1. <span data-ttu-id="47725-176">Выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="47725-176">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="47725-177">После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен) и будет загружена ваша неопубликованная надстройка.</span><span class="sxs-lookup"><span data-stu-id="47725-177">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > <span data-ttu-id="47725-178">Если надстройка не была автоматически загружена, следуйте инструкциям [в Sideload Outlook надстройки для](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) тестирования, чтобы вручную перезагрузить надстройку в Outlook.</span><span class="sxs-lookup"><span data-stu-id="47725-178">If your add-in wasn't automatically sideloaded, then follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) to manually sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="47725-179">Создайте новое сообщение в веб-версии Outlook.</span><span class="sxs-lookup"><span data-stu-id="47725-179">In Outlook on the web, create a new message.</span></span>

    ![Скриншот окна сообщения в Outlook в Интернете с темой, установленной на compose](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="47725-181">В Outlook на Windows, создать новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="47725-181">In Outlook on Windows, create a new message.</span></span>

    ![Скриншот окна сообщения в Outlook на Windows с темой, установленной на compose](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > <span data-ttu-id="47725-183">Если вы работаете надстройок от localhost и видите ошибку "Простите, мы не могли получить доступ *к "ваш-добавить в имя-здесь"*.</span><span class="sxs-lookup"><span data-stu-id="47725-183">If you're running your add-in from localhost and see the error "We're sorry, we couldn't access *{your-add-in-name-here}*.</span></span> <span data-ttu-id="47725-184">Убедитесь, что у вас есть сетевое соединение.</span><span class="sxs-lookup"><span data-stu-id="47725-184">Make sure you have a network connection.</span></span> <span data-ttu-id="47725-185">Если проблема продолжается, пожалуйста, повторите попытку позже.", Возможно, потребуется включить исключение из цикла.</span><span class="sxs-lookup"><span data-stu-id="47725-185">If the problem continues, please try again later.", you may need to enable a loopback exemption.</span></span>
    >
    > 1. <span data-ttu-id="47725-186">Закройте Outlook.</span><span class="sxs-lookup"><span data-stu-id="47725-186">Close Outlook.</span></span>
    > 1. <span data-ttu-id="47725-187">Откройте менеджера **задач и** убедитесь, **чтоmsoadfsb.exe** процесс не работает.</span><span class="sxs-lookup"><span data-stu-id="47725-187">Open the **Task Manager** and ensure that the **msoadfsb.exe** process is not running.</span></span>
    > 1. <span data-ttu-id="47725-188">Выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="47725-188">Run the following command.</span></span>
    >
    >    ```command&nbsp;line
    >    call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >    ```
    >
    > 1. <span data-ttu-id="47725-189">Перезапустите Outlook.</span><span class="sxs-lookup"><span data-stu-id="47725-189">Restart Outlook.</span></span>

## <a name="debug"></a><span data-ttu-id="47725-190">Debug</span><span class="sxs-lookup"><span data-stu-id="47725-190">Debug</span></span>

<span data-ttu-id="47725-191">При внесении изменений в обработку событий запуска в надстройку следует знать, что:</span><span class="sxs-lookup"><span data-stu-id="47725-191">As you make changes to launch-event handling in your add-in, you should be aware that:</span></span>

- <span data-ttu-id="47725-192">Если вы обновили манифест, [удалите надстройку, а](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) затем загрузите ее снова.</span><span class="sxs-lookup"><span data-stu-id="47725-192">If you updated the manifest, [remove the add-in](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) then sideload it again.</span></span>
- <span data-ttu-id="47725-193">Если вы внесли изменения в файлы, не в которые был манифест, закройте и Outlook на Windows или обновите вкладку браузера, Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="47725-193">If you made changes to files other than the manifest, close and reopen Outlook on Windows, or refresh the browser tab running Outlook on the web.</span></span>

<span data-ttu-id="47725-194">При реализации собственной функциональности может потребоваться отладить свой код.</span><span class="sxs-lookup"><span data-stu-id="47725-194">While implementing your own functionality, you may need to debug your code.</span></span> <span data-ttu-id="47725-195">Для получения рекомендаций о том, как отладить активацию надстройки на [Outlook](debug-autolaunch.md)основе событий, см.</span><span class="sxs-lookup"><span data-stu-id="47725-195">For guidance on how to debug event-based add-in activation, see [Debug your event-based Outlook add-in](debug-autolaunch.md).</span></span>

<span data-ttu-id="47725-196">Запись времени выполнения также доступна для этой функции на Windows.</span><span class="sxs-lookup"><span data-stu-id="47725-196">Runtime logging is also available for this feature on Windows.</span></span> <span data-ttu-id="47725-197">Для получения дополнительной информации [см.](../testing/runtime-logging.md#runtime-logging-on-windows)</span><span class="sxs-lookup"><span data-stu-id="47725-197">For more information, see [Debug your add-in with runtime logging](../testing/runtime-logging.md#runtime-logging-on-windows).</span></span>

## <a name="deploy-to-users"></a><span data-ttu-id="47725-198">Развертывание для пользователей</span><span class="sxs-lookup"><span data-stu-id="47725-198">Deploy to users</span></span>

<span data-ttu-id="47725-199">Вы можете развернуть дополнения на основе событий, загрузив манифест через Microsoft 365 администратора.</span><span class="sxs-lookup"><span data-stu-id="47725-199">You can deploy event-based add-ins by uploading the manifest through the Microsoft 365 admin center.</span></span> <span data-ttu-id="47725-200">На портале администратора расширьте раздел **Параметры навигационном** стекле, а затем выберите **интегрированные приложения.**</span><span class="sxs-lookup"><span data-stu-id="47725-200">In the admin portal, expand the **Settings** section in the navigation pane then select **Integrated apps**.</span></span> <span data-ttu-id="47725-201">На странице **Интегрированные приложения** выберите Upload **пользовательских приложений.**</span><span class="sxs-lookup"><span data-stu-id="47725-201">On the **Integrated apps** page, choose the **Upload custom apps** action.</span></span>

![Скриншот страницы интегрированных приложений в центре Microsoft 365, включая Upload пользовательских приложений](../images/outlook-deploy-event-based-add-ins.png)

<span data-ttu-id="47725-203">AppSource и магазины inclient: Возможность развертывания надстройок на основе событий или обновления существующих надстройок, включая функцию активации на основе событий, должна быть доступна в ближайшее время.</span><span class="sxs-lookup"><span data-stu-id="47725-203">AppSource and inclient stores: The ability to deploy event-based add-ins or update existing add-ins to include the event-based activation feature should be available soon.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="47725-204">Надстройки на основе событий ограничиваются только развертыванием, управляемым администратором.</span><span class="sxs-lookup"><span data-stu-id="47725-204">Event-based add-ins are restricted to admin-managed deployments only.</span></span> <span data-ttu-id="47725-205">На данный момент пользователи не могут получить дополнения на основе событий из AppSource или inclient магазинов.</span><span class="sxs-lookup"><span data-stu-id="47725-205">For now, users can't get event-based add-ins from AppSource or inclient stores.</span></span>

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="47725-206">Поведение активации на основе событий и ограничения</span><span class="sxs-lookup"><span data-stu-id="47725-206">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="47725-207">Ожидается, что обработчики дополнительных событий будут короткими, легкими и неинвазивными.</span><span class="sxs-lookup"><span data-stu-id="47725-207">Add-in launch-event handlers are expected to be short-running, lightweight, and as noninvasive as possible.</span></span> <span data-ttu-id="47725-208">После активации надстройки будут тайм-аут в течение примерно 300 секунд, максимальное время, разрешенное для запуска надстройок на основе событий. Чтобы сигнализировать о том, что надстройки завершили обработку события запуска, мы рекомендуем вам позвонить в метод связанному `event.completed` обработчику.</span><span class="sxs-lookup"><span data-stu-id="47725-208">After activation, your add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. To signal that your add-in has completed processing a launch event, we recommend you have the associated handler call the `event.completed` method.</span></span> <span data-ttu-id="47725-209">(Обратите внимание, что код, `event.completed` включенный после выписки, не гарантируется запуск.) Каждый раз, когда срабатывает событие, срабатываемое с ручками надстройок, надстройка активируется и запускается связанный обработчик событий, а окно тайм-аута сбрасывается.</span><span class="sxs-lookup"><span data-stu-id="47725-209">(Note that code included after the `event.completed` statement is not guaranteed to run.) Each time an event that your add-in handles is triggered, the add-in is reactivated and runs the associated event handler, and the timeout window is reset.</span></span> <span data-ttu-id="47725-210">Надстройка заканчивается после того, как она раз, или пользователь закрывает окно compose или отправляет элемент.</span><span class="sxs-lookup"><span data-stu-id="47725-210">The add-in ends after it times out, or the user closes the compose window or sends the item.</span></span>

<span data-ttu-id="47725-211">Если пользователь имеет несколько надстройок, подписавшихся на одно и то же событие, Outlook запускает надстройки в определенном порядке.</span><span class="sxs-lookup"><span data-stu-id="47725-211">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="47725-212">В настоящее время только пять надстройок на основе событий могут активно работать.</span><span class="sxs-lookup"><span data-stu-id="47725-212">Currently, only five event-based add-ins can be actively running.</span></span>

<span data-ttu-id="47725-213">Пользователь может переключиться или перейти от текущего элемента почты, где началось запуск надстройка.</span><span class="sxs-lookup"><span data-stu-id="47725-213">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="47725-214">Запущенная надстройа завершит свою работу в фоновом режиме.</span><span class="sxs-lookup"><span data-stu-id="47725-214">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="47725-215">Некоторые Office.js API, которые изменяют или изменяют пользовательский интерфейс, не допускаются из надстройок на основе событий. Ниже приведены заблокированные API:</span><span class="sxs-lookup"><span data-stu-id="47725-215">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="47725-216">В соответствии `OfficeRuntime.auth` с :</span><span class="sxs-lookup"><span data-stu-id="47725-216">Under `OfficeRuntime.auth`:</span></span>
  - <span data-ttu-id="47725-217">`getAccessToken`(Windows только)</span><span class="sxs-lookup"><span data-stu-id="47725-217">`getAccessToken` (Windows only)</span></span>
- <span data-ttu-id="47725-218">В соответствии `Office.context.auth` с :</span><span class="sxs-lookup"><span data-stu-id="47725-218">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`
- <span data-ttu-id="47725-219">В соответствии `Office.context.mailbox` с :</span><span class="sxs-lookup"><span data-stu-id="47725-219">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="47725-220">В соответствии `Office.context.mailbox.item` с :</span><span class="sxs-lookup"><span data-stu-id="47725-220">Under `Office.context.mailbox.item`:</span></span>
  - `close`
- <span data-ttu-id="47725-221">В соответствии `Office.context.ui` с :</span><span class="sxs-lookup"><span data-stu-id="47725-221">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a><span data-ttu-id="47725-222">См. также</span><span class="sxs-lookup"><span data-stu-id="47725-222">See also</span></span>

- [<span data-ttu-id="47725-223">Манифесты надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="47725-223">Outlook add-in manifests</span></span>](manifests.md)
- [<span data-ttu-id="47725-224">Как отладить дополнения на основе событий</span><span class="sxs-lookup"><span data-stu-id="47725-224">How to debug event-based add-ins</span></span>](debug-autolaunch.md)
