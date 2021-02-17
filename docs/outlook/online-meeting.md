---
title: Создание мобильной надстройки Outlook для поставщика собраний по сети
description: В этой теме обсуждается настройка мобильной надстройки Outlook для поставщика услуг собраний по сети.
ms.topic: article
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: fb98ddeeef8615476659a0abb798ea7901d81248
ms.sourcegitcommit: 1cdf5728102424a46998e1527508b4e7f9f74a4c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/17/2021
ms.locfileid: "50270744"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider"></a><span data-ttu-id="a60ca-103">Создание мобильной надстройки Outlook для поставщика собраний по сети</span><span class="sxs-lookup"><span data-stu-id="a60ca-103">Create an Outlook mobile add-in for an online-meeting provider</span></span>

<span data-ttu-id="a60ca-104">Настройка собрания по сети — это основная задача пользователя Outlook, и ее легко создать с помощью [Outlook](/microsoftteams/teams-add-in-for-outlook) Mobile.</span><span class="sxs-lookup"><span data-stu-id="a60ca-104">Setting up an online meeting is a core experience for an Outlook user, and it's easy to [create a Teams meeting with Outlook](/microsoftteams/teams-add-in-for-outlook) mobile.</span></span> <span data-ttu-id="a60ca-105">Однако создание собрания по сети в Outlook с помощью службы, не относякой к Майкрософт, может быть очень важным.</span><span class="sxs-lookup"><span data-stu-id="a60ca-105">However, creating an online meeting in Outlook with a non-Microsoft service can be cumbersome.</span></span> <span data-ttu-id="a60ca-106">Реализуя эту функцию, поставщики услуг могут упростить создание собраний по сети для пользователей надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="a60ca-106">By implementing this feature, service providers can streamline the online meeting creation experience for their Outlook add-in users.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a60ca-107">Эта функция поддерживается только на Android и iOS с подпиской на Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="a60ca-107">This feature is only supported on Android and iOS with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="a60ca-108">В этой статье вы узнаете, как настроить мобильную надстройка Outlook, чтобы пользователи могли организовывать собрания и присоединяться к ним с помощью службы собраний по сети.</span><span class="sxs-lookup"><span data-stu-id="a60ca-108">In this article, you'll learn how to set up your Outlook mobile add-in to enable users to organize and join a meeting using your online-meeting service.</span></span> <span data-ttu-id="a60ca-109">В этой статье мы будем использовать вымышленного поставщика услуг онлайн-собраний Contoso.</span><span class="sxs-lookup"><span data-stu-id="a60ca-109">Throughout this article, we'll be using a fictional online-meeting service provider, "Contoso".</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="a60ca-110">Настройка среды</span><span class="sxs-lookup"><span data-stu-id="a60ca-110">Set up your environment</span></span>

<span data-ttu-id="a60ca-111">Завершите [краткое начало работы с Outlook,](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) которое создает проект надстройки с помощью генератора Yeoman для надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="a60ca-111">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="a60ca-112">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="a60ca-112">Configure the manifest</span></span>

<span data-ttu-id="a60ca-113">Чтобы позволить пользователям создавать собрания по сети с помощью надстройки, необходимо настроить точку расширения [MobileOnlineMeetingCommandSurface](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface) в манифесте в родительском `MobileFormFactor` элементе.</span><span class="sxs-lookup"><span data-stu-id="a60ca-113">To enable users to create online meetings with your add-in, you must configure the [MobileOnlineMeetingCommandSurface extension point](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface) in the manifest under the parent element `MobileFormFactor`.</span></span> <span data-ttu-id="a60ca-114">Другие форм-факторы не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="a60ca-114">Other form factors are not supported.</span></span>

1. <span data-ttu-id="a60ca-115">В редакторе кода откройте проект быстрого запуска.</span><span class="sxs-lookup"><span data-stu-id="a60ca-115">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="a60ca-116">Откройте файл **manifest.xml,** расположенный в корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="a60ca-116">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="a60ca-117">Выберите весь узел (включая открытые и закрываемые `<VersionOverrides>` теги) и замените его на следующий XML-</span><span class="sxs-lookup"><span data-stu-id="a60ca-117">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Description resid="residDescription"></Description>
    <Requirements>
      <bt:Sets>
        <bt:Set Name="Mailbox" MinVersion="1.3"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="residFunctionFile"/>
          <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="apptComposeGroup">
                <Label resid="residDescription"/>
                <Control xsi:type="Button" id="insertMeetingButton">
                  <Label resid="residLabel"/>
                  <Supertip>
                    <Title resid="residLabel"/>
                    <Description resid="residTooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon-16"/>
                    <bt:Image size="32" resid="icon-32"/>
                    <bt:Image size="64" resid="icon-64"/>
                    <bt:Image size="80" resid="icon-80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>insertContosoMeeting</FunctionName>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>

        <MobileFormFactor>
          <FunctionFile resid="residFunctionFile"/>
          <ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
            <Control xsi:type="MobileButton" id="insertMeetingButton">
              <Label resid="residLabel"/>
              <Icon>
                <bt:Image size="25" scale="1" resid="icon-16"/>
                <bt:Image size="25" scale="2" resid="icon-16"/>
                <bt:Image size="25" scale="3" resid="icon-16"/>

                <bt:Image size="32" scale="1" resid="icon-32"/>
                <bt:Image size="32" scale="2" resid="icon-32"/>
                <bt:Image size="32" scale="3" resid="icon-32"/>

                <bt:Image size="48" scale="1" resid="icon-48"/>
                <bt:Image size="48" scale="2" resid="icon-48"/>
                <bt:Image size="48" scale="3" resid="icon-48"/>
              </Icon>
              <Action xsi:type="ExecuteFunction">
                <FunctionName>insertContosoMeeting</FunctionName>
              </Action>
            </Control>
          </ExtensionPoint>
        </MobileFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="icon-16" DefaultValue="https://contoso.com/assets/icon-16.png"/>
        <bt:Image id="icon-32" DefaultValue="https://contoso.com/assets/icon-32.png"/>
        <bt:Image id="icon-48" DefaultValue="https://contoso.com/assets/icon-48.png"/>
        <bt:Image id="icon-64" DefaultValue="https://contoso.com/assets/icon-64.png"/>
        <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residDescription" DefaultValue="Contoso meeting"/>
        <bt:String id="residLabel" DefaultValue="Add a contoso meeting"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residTooltip" DefaultValue="Add a contoso meeting to this appointment."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

> [!TIP]
> <span data-ttu-id="a60ca-118">Дополнительные информацию о манифестах надстройки Outlook см. в манифестах надстройки [Outlook](manifests.md) и добавлении поддержки команд надстройки [для Outlook Mobile.](add-mobile-support.md)</span><span class="sxs-lookup"><span data-stu-id="a60ca-118">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md) and [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).</span></span>

## <a name="implement-adding-online-meeting-details"></a><span data-ttu-id="a60ca-119">Реализация добавления сведений о собрании по сети</span><span class="sxs-lookup"><span data-stu-id="a60ca-119">Implement adding online meeting details</span></span>

<span data-ttu-id="a60ca-120">В этом разделе вы узнаете, как скрипт надстройки может обновить собрание пользователя, включив сведения о собрании по сети.</span><span class="sxs-lookup"><span data-stu-id="a60ca-120">In this section, learn how your add-in script can update a user's meeting to include online meeting details.</span></span>

1. <span data-ttu-id="a60ca-121">В том же проекте быстрого запуска откройте файл **./src/commands/commands.js** в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="a60ca-121">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="a60ca-122">Замените все содержимое файла **commands.js** следующим javaScript.</span><span class="sxs-lookup"><span data-stu-id="a60ca-122">Replace the entire content of the **commands.js** file with the following JavaScript.</span></span>

    ```js
    // 1. How to construct online meeting details.
    // Not shown: How to get the meeting organizer's ID and other details from your service.
    const newBody = '<br>' +
        '<a href="https://contoso.com/meeting?id=123456789" target="_blank">Join Contoso meeting</a>' +
        '<br><br>' +
        'Phone Dial-in: +1(123)456-7890' +
        '<br><br>' +
        'Meeting ID: 123 456 789' +
        '<br><br>' +
        'Want to test your video connection?' +
        '<br><br>' +
        '<a href="https://contoso.com/testmeeting" target="_blank">Join test meeting</a>' +
        '<br><br>';

    var mailboxItem;

    // Office is ready.
    Office.onReady(function () {
            mailboxItem = Office.context.mailbox.item;
        }
    );

    // 2. How to define a UI-less function named `insertContosoMeeting` (referenced in the manifest)
    //    to update the meeting body with the online meeting details.
    function insertContosoMeeting(event) {
        // Get HTML body from the client.
        mailboxItem.body.getAsync("html",
            { asyncContext: event },
            function (getBodyResult) {
                if (getBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                    updateBody(getBodyResult.asyncContext, getBodyResult.value);
                } else {
                    console.error("Failed to get HTML body.");
                    getBodyResult.asyncContext.completed({ allowEvent: false });
                }
            }
        );
    }

    // 3. How to implement a supporting function `updateBody`
    //    that appends the online meeting details to the current body of the meeting.
    function updateBody(event, existingBody) {
        // Append new body to the existing body.
        mailboxItem.body.setAsync(existingBody + newBody,
            { asyncContext: event, coercionType: "html" },
            function (setBodyResult) {
                if (setBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                    setBodyResult.asyncContext.completed({ allowEvent: true });
                } else {
                    console.error("Failed to set HTML body.");
                    setBodyResult.asyncContext.completed({ allowEvent: false });
                }
            }
        );
    }

    function getGlobal() {
      return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
        ? window
        : typeof global !== "undefined"
        ? global
        : undefined;
    }

    const g = getGlobal();

    // The add-in command functions need to be available in global scope.
    g.insertContosoMeeting = insertContosoMeeting;
    ```

## <a name="testing-and-validation"></a><span data-ttu-id="a60ca-123">Тестирование и проверка</span><span class="sxs-lookup"><span data-stu-id="a60ca-123">Testing and validation</span></span>

<span data-ttu-id="a60ca-124">Следуйте обычным рекомендациям [по проверке и проверке надстройки.](testing-and-tips.md)</span><span class="sxs-lookup"><span data-stu-id="a60ca-124">Follow the usual guidance to [test and validate your add-in](testing-and-tips.md).</span></span> <span data-ttu-id="a60ca-125">После [загрузки](sideload-outlook-add-ins-for-testing.md) неогрузки в Outlook в Интернете, Windows или Mac перезапустите Outlook на мобильном устройстве с Android.</span><span class="sxs-lookup"><span data-stu-id="a60ca-125">After [sideloading](sideload-outlook-add-ins-for-testing.md) in Outlook on the web, Windows, or Mac, restart Outlook on your Android mobile device.</span></span> <span data-ttu-id="a60ca-126">(На данный момент единственным поддерживаемым клиентом является Android.) Затем на новом экране собрания убедитесь, что толль Microsoft Teams или Skype заменен вашим.</span><span class="sxs-lookup"><span data-stu-id="a60ca-126">(Android is the only supported client for now.) Then, on a new meeting screen, verify that the Microsoft Teams or Skype toggle is replaced with your own.</span></span>

### <a name="create-meeting-ui"></a><span data-ttu-id="a60ca-127">Создание пользовательского интерфейса собрания</span><span class="sxs-lookup"><span data-stu-id="a60ca-127">Create meeting UI</span></span>

<span data-ttu-id="a60ca-128">В качестве организатора собрания при создании собрания должны появиться экраны, аналогичные следующим трем изображениям.</span><span class="sxs-lookup"><span data-stu-id="a60ca-128">As a meeting organizer, you should see screens similar to the following three images when you create a meeting.</span></span>

<span data-ttu-id="a60ca-129">[ ![ Screenshot of create meeting screen on Android - Contoso toggle off](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [ ![ screenshot of create meeting screen on Android - loading Contoso toggle](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [ ![ screenshot of create meeting screen on Android - Contoso toggle on](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="a60ca-129">[![screenshot of create meeting screen on Android - Contoso toggle off](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [![screenshot of create meeting screen on Android - loading Contoso toggle](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [![screenshot of create meeting screen on Android - Contoso toggle on](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span></span>

### <a name="join-meeting-ui"></a><span data-ttu-id="a60ca-130">Присоединяйтесь к пользовательскому интерфейсу собрания</span><span class="sxs-lookup"><span data-stu-id="a60ca-130">Join meeting UI</span></span>

<span data-ttu-id="a60ca-131">В качестве участника собрания при просмотре собрания должен отобраться экран, подобный следующему.</span><span class="sxs-lookup"><span data-stu-id="a60ca-131">As a meeting attendee, you should see a screen similar to the following image when you view the meeting.</span></span>

<span data-ttu-id="a60ca-132">[![снимок экрана присоединиться к собранию на Android](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="a60ca-132">[![screenshot of join meeting screen on Android](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a60ca-133">Если вы не видите  ссылку "Присоединиться", возможно, шаблон собрания по сети для вашей службы не зарегистрирован на наших серверах.</span><span class="sxs-lookup"><span data-stu-id="a60ca-133">If you don't see the **Join** link, it may be that the online-meeting template for your service is not registered on our servers.</span></span> <span data-ttu-id="a60ca-134">Подробные сведения см. в разделе "Регистрация [шаблона собрания](#register-your-online-meeting-template) по сети".</span><span class="sxs-lookup"><span data-stu-id="a60ca-134">See the [Register your online-meeting template](#register-your-online-meeting-template) section for details.</span></span>

## <a name="register-your-online-meeting-template"></a><span data-ttu-id="a60ca-135">Регистрация шаблона собрания по сети</span><span class="sxs-lookup"><span data-stu-id="a60ca-135">Register your online-meeting template</span></span>

<span data-ttu-id="a60ca-136">Если вы хотите зарегистрировать шаблон собрания по сети для своей службы, вы можете создать проблему с GitHub с подробными сведениями.</span><span class="sxs-lookup"><span data-stu-id="a60ca-136">If you would like to register the online-meeting template for your service, you can create a GitHub issue with the details.</span></span> <span data-ttu-id="a60ca-137">После этого мы свяемся с вами, чтобы скоординировать временную шкалу регистрации.</span><span class="sxs-lookup"><span data-stu-id="a60ca-137">After that, we'll contact you to coordinate registration timeline.</span></span>

1. <span data-ttu-id="a60ca-138">Перейдите в раздел **"Отзывы"** в конце этой статьи.</span><span class="sxs-lookup"><span data-stu-id="a60ca-138">Go to the **Feedback** section at the end of this article.</span></span>
1. <span data-ttu-id="a60ca-139">Нажмите **ссылку "Эта страница".**</span><span class="sxs-lookup"><span data-stu-id="a60ca-139">Press the **This page** link.</span></span>
1. <span data-ttu-id="a60ca-140">**Задайте для новой** проблемы заголовок "Регистрация шаблона собрания по сети для моей службы", заменив ее `my-service` именем службы.</span><span class="sxs-lookup"><span data-stu-id="a60ca-140">Set the **Title** of the new issue to "Register the online-meeting template for my-service", replacing `my-service` with your service name.</span></span>
1. <span data-ttu-id="a60ca-141">В тексте проблемы замените строку "[Введите здесь отзыв]" строкой, заданной в переменной или аналогичной из раздела "Реализация добавления сведений о собрании по сети" ранее `newBody` в этой статье. [](#implement-adding-online-meeting-details)</span><span class="sxs-lookup"><span data-stu-id="a60ca-141">In the issue body, replace the string "[Enter feedback here]" with the string you set in the `newBody` or similar variable from the [Implement adding online meeting details](#implement-adding-online-meeting-details) section earlier in this article.</span></span>
1. <span data-ttu-id="a60ca-142">Нажмите **кнопку "Отправить новую проблему"**.</span><span class="sxs-lookup"><span data-stu-id="a60ca-142">Click **Submit new issue**.</span></span>

![снимок экрана с новым экраном проблемы GitHub с образцом контента Contoso](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a><span data-ttu-id="a60ca-144">Доступные API</span><span class="sxs-lookup"><span data-stu-id="a60ca-144">Available APIs</span></span>

<span data-ttu-id="a60ca-145">Для этой функции доступны следующие API.</span><span class="sxs-lookup"><span data-stu-id="a60ca-145">The following APIs are available for this feature.</span></span>

- <span data-ttu-id="a60ca-146">API организатора встреч</span><span class="sxs-lookup"><span data-stu-id="a60ca-146">Appointment Organizer APIs</span></span>
  - <span data-ttu-id="a60ca-147">[Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) ([Subject)](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="a60ca-147">[Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="a60ca-148">[Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) ([Time)](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="a60ca-148">[Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="a60ca-149">[Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="a60ca-149">[Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="a60ca-150">[Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) ([Location)](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="a60ca-150">[Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="a60ca-151">[Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalattendees) ([Recipients)](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="a60ca-151">[Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="a60ca-152">[Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredattendees) ([Recipients)](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="a60ca-152">[Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="a60ca-153">[Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) ([Body.getAsync,](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#getasync-coerciontype--options--callback-) [Body.setAsync)](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setasync-data--options--callback-)</span><span class="sxs-lookup"><span data-stu-id="a60ca-153">[Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#getasync-coerciontype--options--callback-), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setasync-data--options--callback-))</span></span>
  - <span data-ttu-id="a60ca-154">[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties)](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="a60ca-154">[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))</span></span>
  - <span data-ttu-id="a60ca-155">[Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))</span><span class="sxs-lookup"><span data-stu-id="a60ca-155">[Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))</span></span>
- <span data-ttu-id="a60ca-156">Обработка потока auth</span><span class="sxs-lookup"><span data-stu-id="a60ca-156">Handle auth flow</span></span>
  - [<span data-ttu-id="a60ca-157">API диалоговых окон</span><span class="sxs-lookup"><span data-stu-id="a60ca-157">Dialog APIs</span></span>](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a><span data-ttu-id="a60ca-158">Ограничения</span><span class="sxs-lookup"><span data-stu-id="a60ca-158">Restrictions</span></span>

<span data-ttu-id="a60ca-159">Применяется несколько ограничений.</span><span class="sxs-lookup"><span data-stu-id="a60ca-159">Several restrictions apply.</span></span>

- <span data-ttu-id="a60ca-160">Применимо только к поставщикам услуг собраний по сети.</span><span class="sxs-lookup"><span data-stu-id="a60ca-160">Applicable only to online-meeting service providers.</span></span>
- <span data-ttu-id="a60ca-161">На экране составить собрание будут отображаться только установленные администратором надстройки, заменяющие параметр Teams или Skype по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="a60ca-161">Only admin-installed add-ins will appear on the meeting compose screen, replacing the default Teams or Skype option.</span></span> <span data-ttu-id="a60ca-162">Установленные пользователем надстройки не активируются.</span><span class="sxs-lookup"><span data-stu-id="a60ca-162">User-installed add-ins won't activate.</span></span>
- <span data-ttu-id="a60ca-163">Значок надстройки должен быть в серой области с использованием hex-кода или его эквивалента `#919191` в [других форматах цвета.](https://convertingcolors.com/hex-color-919191.html)</span><span class="sxs-lookup"><span data-stu-id="a60ca-163">The add-in icon should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>
- <span data-ttu-id="a60ca-164">В режиме организатора встреч (составить) поддерживается только одна команда без пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="a60ca-164">Only one UI-less command is supported in Appointment Organizer (compose) mode.</span></span>

## <a name="see-also"></a><span data-ttu-id="a60ca-165">См. также</span><span class="sxs-lookup"><span data-stu-id="a60ca-165">See also</span></span>

- [<span data-ttu-id="a60ca-166">Надстройки для Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="a60ca-166">Add-ins for Outlook Mobile</span></span>](outlook-mobile-addins.md)
- [<span data-ttu-id="a60ca-167">Добавление поддержки команд надстройки для Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="a60ca-167">Add support for add-in commands for Outlook Mobile</span></span>](add-mobile-support.md)
