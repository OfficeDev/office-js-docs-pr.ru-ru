---
title: Настройка надстройки Outlook для активации на основе событий
description: Узнайте, как настроить Outlook надстройку для активации на основе событий.
ms.topic: article
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 1856f78b7e6d49952d2eebf521894d6a988402a0
ms.sourcegitcommit: 30a861ece18255e342725e31c47f01960b854532
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/16/2021
ms.locfileid: "53455532"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation"></a><span data-ttu-id="ee61a-103">Настройка надстройки Outlook для активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="ee61a-103">Configure your Outlook add-in for event-based activation</span></span>

<span data-ttu-id="ee61a-104">Без функции активации на основе событий пользователю необходимо явно запустить надстройки для выполнения задач.</span><span class="sxs-lookup"><span data-stu-id="ee61a-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="ee61a-105">Эта функция позволяет надстройки выполнять задачи на основе определенных событий, особенно для операций, применимых к каждому элементу.</span><span class="sxs-lookup"><span data-stu-id="ee61a-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="ee61a-106">Вы также можете интегрироваться с области задач и функциональными возможностями без пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="ee61a-106">You can also integrate with the task pane and UI-less functionality.</span></span>

<span data-ttu-id="ee61a-107">К концу этого погона у вас будет надстройка, которая запускается всякий раз, когда создается новый элемент и задает объект.</span><span class="sxs-lookup"><span data-stu-id="ee61a-107">By the end of this walkthrough, you'll have an add-in that runs whenever a new item is created and sets the subject.</span></span>

> [!NOTE]
> <span data-ttu-id="ee61a-108">Поддержка этой функции была представлена в [наборе требований 1.10](../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md).</span><span class="sxs-lookup"><span data-stu-id="ee61a-108">Support for this feature was introduced in [requirement set 1.10](../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span> <span data-ttu-id="ee61a-109">См [клиенты и платформы](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.</span><span class="sxs-lookup"><span data-stu-id="ee61a-109">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-events"></a><span data-ttu-id="ee61a-110">Поддерживаемые события</span><span class="sxs-lookup"><span data-stu-id="ee61a-110">Supported events</span></span>

<span data-ttu-id="ee61a-111">В настоящее время следующие события поддерживаются в Интернете и Windows.</span><span class="sxs-lookup"><span data-stu-id="ee61a-111">At present, the following events are supported on the web and on Windows.</span></span>

|<span data-ttu-id="ee61a-112">Событие</span><span class="sxs-lookup"><span data-stu-id="ee61a-112">Event</span></span>|<span data-ttu-id="ee61a-113">Описание</span><span class="sxs-lookup"><span data-stu-id="ee61a-113">Description</span></span>|<span data-ttu-id="ee61a-114">Minimum</span><span class="sxs-lookup"><span data-stu-id="ee61a-114">Minimum</span></span><br><span data-ttu-id="ee61a-115">набор требований</span><span class="sxs-lookup"><span data-stu-id="ee61a-115">requirement set</span></span>|
|---|---|---|
|`OnNewMessageCompose`|<span data-ttu-id="ee61a-116">При составлении нового сообщения (включает ответ, ответ все и вперед), но не при редактировании, например, черновика.</span><span class="sxs-lookup"><span data-stu-id="ee61a-116">On composing a new message (includes reply, reply all, and forward) but not on editing, for example, a draft.</span></span>|<span data-ttu-id="ee61a-117">1.10</span><span class="sxs-lookup"><span data-stu-id="ee61a-117">1.10</span></span>|
|`OnNewAppointmentOrganizer`|<span data-ttu-id="ee61a-118">О создании новой встречи, но не о редактировании существующего.</span><span class="sxs-lookup"><span data-stu-id="ee61a-118">On creating a new appointment but not on editing an existing one.</span></span>|<span data-ttu-id="ee61a-119">1.10</span><span class="sxs-lookup"><span data-stu-id="ee61a-119">1.10</span></span>|
|`OnMessageAttachmentsChanged`|<span data-ttu-id="ee61a-120">При добавлении или удалении вложений при сочинении сообщения.</span><span class="sxs-lookup"><span data-stu-id="ee61a-120">On adding or removing attachments while composing a message.</span></span>|<span data-ttu-id="ee61a-121">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="ee61a-121">Preview</span></span>|
|`OnAppointmentAttachmentsChanged`|<span data-ttu-id="ee61a-122">При добавлении или удалении вложений при записи на прием.</span><span class="sxs-lookup"><span data-stu-id="ee61a-122">On adding or removing attachments while composing an appointment.</span></span>|<span data-ttu-id="ee61a-123">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="ee61a-123">Preview</span></span>|
|`OnMessageRecipientsChanged`|<span data-ttu-id="ee61a-124">При добавлении или удалении получателей при сочинении сообщения.</span><span class="sxs-lookup"><span data-stu-id="ee61a-124">On adding or removing recipients while composing a message.</span></span>|<span data-ttu-id="ee61a-125">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="ee61a-125">Preview</span></span>|
|`OnAppointmentAttendeesChanged`|<span data-ttu-id="ee61a-126">При добавлении или удалении участников при записи на прием.</span><span class="sxs-lookup"><span data-stu-id="ee61a-126">On adding or removing attendees while composing an appointment.</span></span>|<span data-ttu-id="ee61a-127">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="ee61a-127">Preview</span></span>|
|`OnAppointmentTimeChanged`|<span data-ttu-id="ee61a-128">При изменении даты и времени при записи на прием.</span><span class="sxs-lookup"><span data-stu-id="ee61a-128">On changing date/time while composing an appointment.</span></span>|<span data-ttu-id="ee61a-129">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="ee61a-129">Preview</span></span>|
|`OnAppointmentRecurrenceChanged`|<span data-ttu-id="ee61a-130">При добавлении, изменении или удалении сведений о повторении при записи на прием.</span><span class="sxs-lookup"><span data-stu-id="ee61a-130">On adding, changing, or removing the recurrence details while composing an appointment.</span></span> <span data-ttu-id="ee61a-131">Если дата и время изменены, `OnAppointmentTimeChanged` событие также будет уволено.</span><span class="sxs-lookup"><span data-stu-id="ee61a-131">If the date/time is changed, the `OnAppointmentTimeChanged` event will also be fired.</span></span>|<span data-ttu-id="ee61a-132">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="ee61a-132">Preview</span></span>|
|`OnInfoBarDismissClicked`|<span data-ttu-id="ee61a-133">При отклонении уведомления при записи сообщения или элемента встречи.</span><span class="sxs-lookup"><span data-stu-id="ee61a-133">On dismissing a notification while composing a message or appointment item.</span></span> <span data-ttu-id="ee61a-134">Уведомления будут получать только надстройка, которая добавила уведомление.</span><span class="sxs-lookup"><span data-stu-id="ee61a-134">Only the add-in that added the notification will be notified.</span></span>|<span data-ttu-id="ee61a-135">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="ee61a-135">Preview</span></span>|

> [!IMPORTANT]
> <span data-ttu-id="ee61a-136">События, которые по-прежнему находятся в предварительном просмотре, доступны только с Microsoft 365 подпиской в Outlook в Интернете и Windows.</span><span class="sxs-lookup"><span data-stu-id="ee61a-136">Events still in preview are only available with a Microsoft 365 subscription in Outlook on the web and on Windows.</span></span> <span data-ttu-id="ee61a-137">Дополнительные сведения см. [в статье How to preview](#how-to-preview) in this article.</span><span class="sxs-lookup"><span data-stu-id="ee61a-137">For more details, see [How to preview](#how-to-preview) in this article.</span></span> <span data-ttu-id="ee61a-138">События предварительного просмотра не следует использовать в производственных надстройках.</span><span class="sxs-lookup"><span data-stu-id="ee61a-138">Preview events shouldn't be used in production add-ins.</span></span>

### <a name="how-to-preview"></a><span data-ttu-id="ee61a-139">Предварительный просмотр</span><span class="sxs-lookup"><span data-stu-id="ee61a-139">How to preview</span></span>

<span data-ttu-id="ee61a-140">Мы приглашаем вас попробовать события в предварительном просмотре!</span><span class="sxs-lookup"><span data-stu-id="ee61a-140">We invite you to try out the events now in preview!</span></span> <span data-ttu-id="ee61a-141">Дайте нам знать о ваших сценариях и о том, как мы можем улучшить ситуацию, GitHub с помощью GitHub (см. раздел **Обратная** связь в конце этой страницы).</span><span class="sxs-lookup"><span data-stu-id="ee61a-141">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="ee61a-142">Чтобы просмотреть эти события:</span><span class="sxs-lookup"><span data-stu-id="ee61a-142">To preview these events:</span></span>

- <span data-ttu-id="ee61a-143">Для Outlook в Интернете:</span><span class="sxs-lookup"><span data-stu-id="ee61a-143">For Outlook on the web:</span></span>
  - <span data-ttu-id="ee61a-144">[Настройка целевого выпуска для](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)Microsoft 365 клиента.</span><span class="sxs-lookup"><span data-stu-id="ee61a-144">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="ee61a-145">Ссылка  на бета-библиотеку на CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="ee61a-145">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="ee61a-146">Файл [определения типа](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) для компиляции и IntelliSense typeScript CDN и [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="ee61a-146">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="ee61a-147">Эти типы можно установить с `npm install --save-dev @types/office-js-preview` помощью .</span><span class="sxs-lookup"><span data-stu-id="ee61a-147">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="ee61a-148">Для Outlook на Windows:</span><span class="sxs-lookup"><span data-stu-id="ee61a-148">For Outlook on Windows:</span></span>
  - <span data-ttu-id="ee61a-149">Минимальная требуемая сборка — 16.0.14026.20000.</span><span class="sxs-lookup"><span data-stu-id="ee61a-149">The minimum required build is 16.0.14026.20000.</span></span> <span data-ttu-id="ee61a-150">Присоединяйтесь [к Office программы insider](https://insider.office.com) для доступа к Office бета-сборки.</span><span class="sxs-lookup"><span data-stu-id="ee61a-150">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>
  - <span data-ttu-id="ee61a-151">Настройка реестра.</span><span class="sxs-lookup"><span data-stu-id="ee61a-151">Configure the registry.</span></span> <span data-ttu-id="ee61a-152">Outlook включает локализованную копию выпуска и бета-версии Office.js вместо загрузки из CDN.</span><span class="sxs-lookup"><span data-stu-id="ee61a-152">Outlook includes a local copy of the production and beta versions of Office.js instead of loading from the CDN.</span></span> <span data-ttu-id="ee61a-153">По умолчанию ссылается локализованная производственная копия API.</span><span class="sxs-lookup"><span data-stu-id="ee61a-153">By default, the local production copy of the API is referenced.</span></span> <span data-ttu-id="ee61a-154">Чтобы перейти на локализованную бета-версию API Outlook JavaScript, необходимо добавить эту запись реестра, в противном случае бета-API не могут быть найдены.</span><span class="sxs-lookup"><span data-stu-id="ee61a-154">To switch to the local beta copy of the Outlook JavaScript APIs, you need to add this registry entry, otherwise beta APIs may not be found.</span></span>
    1. <span data-ttu-id="ee61a-155">Создание ключа `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer` реестра.</span><span class="sxs-lookup"><span data-stu-id="ee61a-155">Create the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer`.</span></span>
    1. <span data-ttu-id="ee61a-156">Добавьте запись с `EnableBetaAPIsInJavaScript` именем и установите значение `1` .</span><span class="sxs-lookup"><span data-stu-id="ee61a-156">Add an entry named `EnableBetaAPIsInJavaScript` and set the value to `1`.</span></span> <span data-ttu-id="ee61a-157">На приведенном ниже изображении показано, как должен выглядеть реестр.</span><span class="sxs-lookup"><span data-stu-id="ee61a-157">The following image shows what the registry should look like.</span></span>

        ![Снимок экрана редактора реестра с ключевым значением реестра EnableBetaAPIsInJavaScript.](../images/outlook-beta-registry-key.png)

## <a name="set-up-your-environment"></a><span data-ttu-id="ee61a-159">Настройка среды</span><span class="sxs-lookup"><span data-stu-id="ee61a-159">Set up your environment</span></span>

<span data-ttu-id="ee61a-160">Выполните [Outlook,](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) который создает проект надстройки с генератором Yeoman для Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="ee61a-160">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="ee61a-161">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="ee61a-161">Configure the manifest</span></span>

<span data-ttu-id="ee61a-162">Чтобы включить активацию надстройки на основе событий, необходимо настроить элемент [Runtimes](../reference/manifest/runtimes.md) и точку расширения [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent) в узле `VersionOverridesV1_1` манифеста.</span><span class="sxs-lookup"><span data-stu-id="ee61a-162">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="ee61a-163">Пока это `DesktopFormFactor` единственный поддерживаемый форм-фактор.</span><span class="sxs-lookup"><span data-stu-id="ee61a-163">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="ee61a-164">В редакторе кода откройте проект быстрого запуска.</span><span class="sxs-lookup"><span data-stu-id="ee61a-164">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="ee61a-165">Откройте **manifest.xml** файл, расположенный в корне проекта.</span><span class="sxs-lookup"><span data-stu-id="ee61a-165">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="ee61a-166">Выберите весь узел (включая открытые и закрываемые теги) и замените его на следующий XML, а затем `<VersionOverrides>` сохраните изменения.</span><span class="sxs-lookup"><span data-stu-id="ee61a-166">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.</span></span>

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

<span data-ttu-id="ee61a-167">Outlook на Windows использует файл JavaScript, а Outlook в Интернете использует HTML-файл, который может ссылаться на тот же файл JavaScript.</span><span class="sxs-lookup"><span data-stu-id="ee61a-167">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="ee61a-168">Необходимо предоставить ссылки на оба этих файла в узле манифеста, так как платформа Outlook в конечном счете определяет, следует ли использовать HTML или JavaScript на основе Outlook `Resources` клиента.</span><span class="sxs-lookup"><span data-stu-id="ee61a-168">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="ee61a-169">Таким образом, чтобы настроить обработку событий, укадь расположение HTML в элементе, а затем в его детском элементе укаймляй расположение файла JavaScript, вписаного или ссылаемого `Runtime` `Override` HTML.</span><span class="sxs-lookup"><span data-stu-id="ee61a-169">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="ee61a-170">Дополнительные информацию о манифестах для Outlook надстройки см. в Outlook [манифестах надстройки.](manifests.md)</span><span class="sxs-lookup"><span data-stu-id="ee61a-170">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="ee61a-171">Реализация обработки событий</span><span class="sxs-lookup"><span data-stu-id="ee61a-171">Implement event handling</span></span>

<span data-ttu-id="ee61a-172">Для выбранных событий необходимо реализовать обработку.</span><span class="sxs-lookup"><span data-stu-id="ee61a-172">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="ee61a-173">В этом сценарии вы добавим обработку для составления новых элементов.</span><span class="sxs-lookup"><span data-stu-id="ee61a-173">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="ee61a-174">В том же проекте быстрого запуска откройте **файл ./src/commands/commands.js** в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="ee61a-174">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="ee61a-175">После `action` функции вставьте следующие функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="ee61a-175">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="ee61a-176">Добавьте следующий код JavaScript в конце файла.</span><span class="sxs-lookup"><span data-stu-id="ee61a-176">Add the following JavaScript code at the end of the file.</span></span>

    ```js
    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
    Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    ```

1. <span data-ttu-id="ee61a-177">Сохраните изменения.</span><span class="sxs-lookup"><span data-stu-id="ee61a-177">Save your changes.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ee61a-178">Windows. В настоящее время импорт не поддерживается в файле JavaScript, где выполняется обработка активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="ee61a-178">Windows: At present, imports are not supported in the JavaScript file where you implement the handling for event-based activation.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="ee61a-179">Проверка</span><span class="sxs-lookup"><span data-stu-id="ee61a-179">Try it out</span></span>

1. <span data-ttu-id="ee61a-180">Выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="ee61a-180">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="ee61a-181">После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен) и будет загружена ваша неопубликованная надстройка.</span><span class="sxs-lookup"><span data-stu-id="ee61a-181">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > <span data-ttu-id="ee61a-182">Если надстройка не была автоматически загружена, следуйте инструкциям в [Sideload Outlook](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) надстройки для тестирования, чтобы вручную разгрузить надстройку в Outlook.</span><span class="sxs-lookup"><span data-stu-id="ee61a-182">If your add-in wasn't automatically sideloaded, then follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) to manually sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="ee61a-183">Создайте новое сообщение в веб-версии Outlook.</span><span class="sxs-lookup"><span data-stu-id="ee61a-183">In Outlook on the web, create a new message.</span></span>

    ![Снимок экрана окна сообщения в Outlook в Интернете с набором субъекта на композицию.](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="ee61a-185">В Outlook Windows создайте новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="ee61a-185">In Outlook on Windows, create a new message.</span></span>

    ![Снимок экрана окна сообщения в Outlook на Windows с набором темы на композит.](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > <span data-ttu-id="ee61a-187">Если вы выполняете надстройки из localhost и видите ошибку "К сожалению, мы не могли получить доступ *{your-add-in-name-here}*.</span><span class="sxs-lookup"><span data-stu-id="ee61a-187">If you're running your add-in from localhost and see the error "We're sorry, we couldn't access *{your-add-in-name-here}*.</span></span> <span data-ttu-id="ee61a-188">Убедитесь, что у вас есть сетевое подключение.</span><span class="sxs-lookup"><span data-stu-id="ee61a-188">Make sure you have a network connection.</span></span> <span data-ttu-id="ee61a-189">Если проблема продолжится, попробуйте еще раз.", возможно, потребуется включить освобождение от циклов.</span><span class="sxs-lookup"><span data-stu-id="ee61a-189">If the problem continues, please try again later.", you may need to enable a loopback exemption.</span></span>
    >
    > 1. <span data-ttu-id="ee61a-190">Закройте Outlook.</span><span class="sxs-lookup"><span data-stu-id="ee61a-190">Close Outlook.</span></span>
    > 1. <span data-ttu-id="ee61a-191">Откройте диспетчер **задач** и убедитесь, что **msoadfsb.exe** процесс не запущен.</span><span class="sxs-lookup"><span data-stu-id="ee61a-191">Open the **Task Manager** and ensure that the **msoadfsb.exe** process is not running.</span></span>
    > 1. <span data-ttu-id="ee61a-192">Выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="ee61a-192">Run the following command.</span></span>
    >
    >    ```command&nbsp;line
    >    call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >    ```
    >
    > 1. <span data-ttu-id="ee61a-193">Перезапустите Outlook.</span><span class="sxs-lookup"><span data-stu-id="ee61a-193">Restart Outlook.</span></span>

## <a name="debug"></a><span data-ttu-id="ee61a-194">Debug</span><span class="sxs-lookup"><span data-stu-id="ee61a-194">Debug</span></span>

<span data-ttu-id="ee61a-195">При внесении изменений в обработку событий запуска в надстройку следует помнить, что:</span><span class="sxs-lookup"><span data-stu-id="ee61a-195">As you make changes to launch-event handling in your add-in, you should be aware that:</span></span>

- <span data-ttu-id="ee61a-196">Если вы обновили манифест, [удалите надстройку,](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) а затем снова разгрузите ее.</span><span class="sxs-lookup"><span data-stu-id="ee61a-196">If you updated the manifest, [remove the add-in](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) then sideload it again.</span></span>
- <span data-ttu-id="ee61a-197">Если вы внося изменения в файлы, помимо манифеста, закрой и Outlook на Windows или обновите вкладку браузера, запущенную Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="ee61a-197">If you made changes to files other than the manifest, close and reopen Outlook on Windows, or refresh the browser tab running Outlook on the web.</span></span>

<span data-ttu-id="ee61a-198">При реализации собственных функций может потребоваться отламывка кода.</span><span class="sxs-lookup"><span data-stu-id="ee61a-198">While implementing your own functionality, you may need to debug your code.</span></span> <span data-ttu-id="ee61a-199">Инструкции по отламывить активацию надстройки на основе событий см. в Outlook [событиях.](debug-autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="ee61a-199">For guidance on how to debug event-based add-in activation, see [Debug your event-based Outlook add-in](debug-autolaunch.md).</span></span>

<span data-ttu-id="ee61a-200">Для этой функции также доступна ведение журнала в Windows.</span><span class="sxs-lookup"><span data-stu-id="ee61a-200">Runtime logging is also available for this feature on Windows.</span></span> <span data-ttu-id="ee61a-201">Дополнительные сведения см. в [рубке Отлаговка](../testing/runtime-logging.md#runtime-logging-on-windows)надстройки с ведением журнала времени работы.</span><span class="sxs-lookup"><span data-stu-id="ee61a-201">For more information, see [Debug your add-in with runtime logging](../testing/runtime-logging.md#runtime-logging-on-windows).</span></span>

## <a name="deploy-to-users"></a><span data-ttu-id="ee61a-202">Развертывание для пользователей</span><span class="sxs-lookup"><span data-stu-id="ee61a-202">Deploy to users</span></span>

<span data-ttu-id="ee61a-203">Вы можете развернуть надстройки на основе событий, загрузив манифест через Центр администрирования Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="ee61a-203">You can deploy event-based add-ins by uploading the manifest through the Microsoft 365 admin center.</span></span> <span data-ttu-id="ee61a-204">На портале администрирования расширь **раздел Параметры** области навигации выберите **интегрированные приложения.**</span><span class="sxs-lookup"><span data-stu-id="ee61a-204">In the admin portal, expand the **Settings** section in the navigation pane then select **Integrated apps**.</span></span> <span data-ttu-id="ee61a-205">На странице **Интегрированные приложения** выберите действие **Upload приложений.**</span><span class="sxs-lookup"><span data-stu-id="ee61a-205">On the **Integrated apps** page, choose the **Upload custom apps** action.</span></span>

![Снимок экрана страницы Интегрированные приложения на Центр администрирования Microsoft 365, включая действие Upload приложений.](../images/outlook-deploy-event-based-add-ins.png)

<span data-ttu-id="ee61a-207">AppSource и inclient stores: возможность развертывания надстройок на основе событий или обновления существующих надстройок для включения функции активации на основе событий должна быть доступна в ближайшее время.</span><span class="sxs-lookup"><span data-stu-id="ee61a-207">AppSource and inclient stores: The ability to deploy event-based add-ins or update existing add-ins to include the event-based activation feature should be available soon.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ee61a-208">Надстройки на основе событий ограничиваются только развертыванием, управляемым администратором.</span><span class="sxs-lookup"><span data-stu-id="ee61a-208">Event-based add-ins are restricted to admin-managed deployments only.</span></span> <span data-ttu-id="ee61a-209">Пока пользователи не могут получать надстройки на основе событий из магазинов AppSource или inclient.</span><span class="sxs-lookup"><span data-stu-id="ee61a-209">For now, users can't get event-based add-ins from AppSource or inclient stores.</span></span>

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="ee61a-210">Поведение и ограничения активации на основе событий</span><span class="sxs-lookup"><span data-stu-id="ee61a-210">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="ee61a-211">Обработчики событий запуска надстройки должны быть короткими, легкими и максимально неинвазативными.</span><span class="sxs-lookup"><span data-stu-id="ee61a-211">Add-in launch-event handlers are expected to be short-running, lightweight, and as noninvasive as possible.</span></span> <span data-ttu-id="ee61a-212">После активации надстройка будет отрабатывания в течение примерно 300 секунд— максимального времени, разрешенного для запуска надстроек на основе событий. Чтобы сигнализировать, что ваша надстройка завершила обработку события запуска, рекомендуется вызвать этот метод с помощью связанного `event.completed` обработера.</span><span class="sxs-lookup"><span data-stu-id="ee61a-212">After activation, your add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. To signal that your add-in has completed processing a launch event, we recommend you have the associated handler call the `event.completed` method.</span></span> <span data-ttu-id="ee61a-213">(Обратите внимание, что код, включенный после запуска заявления, `event.completed` не гарантируется.) Каждый раз, когда запускается событие, запускаемое ручками надстройки, надстройка активируется и запускает связанный обработок событий, а окно времени сброшено.</span><span class="sxs-lookup"><span data-stu-id="ee61a-213">(Note that code included after the `event.completed` statement is not guaranteed to run.) Each time an event that your add-in handles is triggered, the add-in is reactivated and runs the associated event handler, and the timeout window is reset.</span></span> <span data-ttu-id="ee61a-214">Надстройка заканчивается после того, как она разовая, или пользователь закрывает окно составить или отправляет элемент.</span><span class="sxs-lookup"><span data-stu-id="ee61a-214">The add-in ends after it times out, or the user closes the compose window or sends the item.</span></span>

<span data-ttu-id="ee61a-215">Если у пользователя есть несколько надстройок, которые подписаны на одно и то же событие, Outlook платформа запускает надстройки без определенного порядка.</span><span class="sxs-lookup"><span data-stu-id="ee61a-215">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="ee61a-216">В настоящее время можно активно запускать только пять надстройок на основе событий.</span><span class="sxs-lookup"><span data-stu-id="ee61a-216">Currently, only five event-based add-ins can be actively running.</span></span>

<span data-ttu-id="ee61a-217">Пользователь может переключаться или перемещаться от текущего элемента почты, где надстройка начала работать.</span><span class="sxs-lookup"><span data-stu-id="ee61a-217">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="ee61a-218">Запущенная надстройка завершит свою работу в фоновом режиме.</span><span class="sxs-lookup"><span data-stu-id="ee61a-218">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="ee61a-219">Импорт не поддерживается в файле JavaScript, где выполняется обработка активации на основе событий в Windows клиенте.</span><span class="sxs-lookup"><span data-stu-id="ee61a-219">Imports are not supported in the JavaScript file where you implement the handling for event-based activation in the Windows client.</span></span>

<span data-ttu-id="ee61a-220">Некоторые Office.js API, которые изменяют или изменяют пользовательский интерфейс, не допускаются из надстройок на основе событий. Ниже заблокировали API.</span><span class="sxs-lookup"><span data-stu-id="ee61a-220">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs.</span></span>

- <span data-ttu-id="ee61a-221">В `OfficeRuntime.auth` статье:</span><span class="sxs-lookup"><span data-stu-id="ee61a-221">Under `OfficeRuntime.auth`:</span></span>
  - <span data-ttu-id="ee61a-222">`getAccessToken`(Windows только)</span><span class="sxs-lookup"><span data-stu-id="ee61a-222">`getAccessToken` (Windows only)</span></span>
- <span data-ttu-id="ee61a-223">В `Office.context.auth` статье:</span><span class="sxs-lookup"><span data-stu-id="ee61a-223">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`
- <span data-ttu-id="ee61a-224">В `Office.context.mailbox` статье:</span><span class="sxs-lookup"><span data-stu-id="ee61a-224">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="ee61a-225">В `Office.context.mailbox.item` статье:</span><span class="sxs-lookup"><span data-stu-id="ee61a-225">Under `Office.context.mailbox.item`:</span></span>
  - `close`
- <span data-ttu-id="ee61a-226">В `Office.context.ui` статье:</span><span class="sxs-lookup"><span data-stu-id="ee61a-226">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`

### <a name="requesting-external-data"></a><span data-ttu-id="ee61a-227">Запрос внешних данных</span><span class="sxs-lookup"><span data-stu-id="ee61a-227">Requesting external data</span></span>

<span data-ttu-id="ee61a-228">Вы можете запрашивать внешние данные с помощью API типа [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) или с помощью [XmlHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest)— стандартного веб-API, который выдает http-запросы для взаимодействия с серверами.</span><span class="sxs-lookup"><span data-stu-id="ee61a-228">You can request external data by using an API like [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="ee61a-229">Следует помнить, что при создании XmlHttpRequests необходимо [](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) использовать дополнительные меры безопасности, требующие одинаковой политики происхождения и [простой CORS.](https://www.w3.org/TR/cors/)</span><span class="sxs-lookup"><span data-stu-id="ee61a-229">Be aware that you must use additional security measures when making XmlHttpRequests, requiring [Same Origin Policy](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="ee61a-230">Простая реализация CORS не может использовать файлы cookie и поддерживает только простые методы (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="ee61a-230">A simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="ee61a-231">Простые запросы CORS принимают простые заголовки с именами полей `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="ee61a-231">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="ee61a-232">Вы также можете использовать `Content-Type` заготку в простой CORS, при условии, что тип контента `application/x-www-form-urlencoded` , `text/plain` или `multipart/form-data` .</span><span class="sxs-lookup"><span data-stu-id="ee61a-232">You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

<span data-ttu-id="ee61a-233">Полная поддержка CORS скоро.</span><span class="sxs-lookup"><span data-stu-id="ee61a-233">Full CORS support is coming soon.</span></span>

## <a name="see-also"></a><span data-ttu-id="ee61a-234">См. также</span><span class="sxs-lookup"><span data-stu-id="ee61a-234">See also</span></span>

- [<span data-ttu-id="ee61a-235">Манифесты надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="ee61a-235">Outlook add-in manifests</span></span>](manifests.md)
- [<span data-ttu-id="ee61a-236">Отламывка надстроек на основе событий</span><span class="sxs-lookup"><span data-stu-id="ee61a-236">How to debug event-based add-ins</span></span>](debug-autolaunch.md)
- [<span data-ttu-id="ee61a-237">Параметры списка AppSource для надстройки на Outlook событий</span><span class="sxs-lookup"><span data-stu-id="ee61a-237">AppSource listing options for your event-based Outlook add-in</span></span>](autolaunch-store-options.md)
- <span data-ttu-id="ee61a-238">Примеры PnP:</span><span class="sxs-lookup"><span data-stu-id="ee61a-238">PnP samples:</span></span>
  - [<span data-ttu-id="ee61a-239">Для Outlook для набора подписи используйте активацию на основе событий</span><span class="sxs-lookup"><span data-stu-id="ee61a-239">Use Outlook event-based activation to set the signature</span></span>](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-set-signature)
  - [<span data-ttu-id="ee61a-240">Использование Outlook активации на основе событий для тегов внешних получателей</span><span class="sxs-lookup"><span data-stu-id="ee61a-240">Use Outlook event-based activation to tag external recipients</span></span>](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-tag-external)
