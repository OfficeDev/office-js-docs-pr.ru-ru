---
title: Создание надстройки Outlook Mobile для поставщика собраний в Интернете
description: Сведения о том, как настроить надстройку Outlook Mobile для поставщика услуг по подключению к интерактивному собранию.
ms.topic: article
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: 9f0b50602ab4941b16c15abe97c3f099a54f5b42
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094003"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider"></a><span data-ttu-id="50193-103">Создание надстройки Outlook Mobile для поставщика собраний в Интернете</span><span class="sxs-lookup"><span data-stu-id="50193-103">Create an Outlook mobile add-in for an online-meeting provider</span></span>

<span data-ttu-id="50193-104">Настройка собрания по сети — это основной интерфейс для пользователя Outlook, который позволяет легко [создать собрание Teams с помощью Outlook](/microsoftteams/teams-add-in-for-outlook) Mobile.</span><span class="sxs-lookup"><span data-stu-id="50193-104">Setting up an online meeting is a core experience for an Outlook user, and it's easy to [create a Teams meeting with Outlook](/microsoftteams/teams-add-in-for-outlook) mobile.</span></span> <span data-ttu-id="50193-105">Однако создание собрания по сети в Outlook со службой, отличной от Майкрософт, может быть утомительным.</span><span class="sxs-lookup"><span data-stu-id="50193-105">However, creating an online meeting in Outlook with a non-Microsoft service can be cumbersome.</span></span> <span data-ttu-id="50193-106">Реализуя эту функцию, поставщики услуг могут упростить процесс создания собраний по сети для пользователей надстроек Outlook.</span><span class="sxs-lookup"><span data-stu-id="50193-106">By implementing this feature, service providers can streamline the online meeting creation experience for their Outlook add-in users.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="50193-107">Эта функция поддерживается только в Android с подпиской на Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="50193-107">This feature is only supported on Android with a Microsoft 365 subscription.</span></span>

<span data-ttu-id="50193-108">В этой статье вы узнаете, как настроить надстройку Outlook Mobile, чтобы позволить пользователям упорядочивать и присоединяться к собранию с помощью службы собраний по сети.</span><span class="sxs-lookup"><span data-stu-id="50193-108">In this article, you'll learn how to set up your Outlook mobile add-in to enable users to organize and join a meeting using your online-meeting service.</span></span> <span data-ttu-id="50193-109">В этой статье мы будем использовать фиктивный поставщик услуг по подключению к собраниям, "contoso".</span><span class="sxs-lookup"><span data-stu-id="50193-109">Throughout this article, we'll be using a fictional online-meeting service provider, "Contoso".</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="50193-110">Настройка среды</span><span class="sxs-lookup"><span data-stu-id="50193-110">Set up your environment</span></span>

<span data-ttu-id="50193-111">Завершите работу с [быстрым запуском Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) , который создает проект надстройки с помощью генератора Yeoman для надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="50193-111">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="50193-112">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="50193-112">Configure the manifest</span></span>

<span data-ttu-id="50193-113">Чтобы позволить пользователям создавать собрания по сети с надстройкой, необходимо настроить `MobileOnlineMeetingCommandSurface` точку расширения в манифесте под родительским элементом `MobileFormFactor` .</span><span class="sxs-lookup"><span data-stu-id="50193-113">To enable users to create online meetings with your add-in, you must configure the `MobileOnlineMeetingCommandSurface` extension point in the manifest under the parent element `MobileFormFactor`.</span></span> <span data-ttu-id="50193-114">Другие конструктивные параметры не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="50193-114">Other form factors are not supported.</span></span>

1. <span data-ttu-id="50193-115">В редакторе кода откройте Быстрый запуск проекта.</span><span class="sxs-lookup"><span data-stu-id="50193-115">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="50193-116">Откройте файл **manifest.xml** , расположенный в корневом каталоге проекта.</span><span class="sxs-lookup"><span data-stu-id="50193-116">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="50193-117">Выберите весь `<VersionOverrides>` узел (включая открывающие и закрывающие теги) и замените его следующим XML-документом.</span><span class="sxs-lookup"><span data-stu-id="50193-117">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
> <span data-ttu-id="50193-118">Чтобы узнать больше о манифестах для надстроек Outlook, ознакомьтесь с разделом [манифесты надстроек Outlook](manifests.md) и [добавьте поддержку команд надстроек для Outlook Mobile](add-mobile-support.md).</span><span class="sxs-lookup"><span data-stu-id="50193-118">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md) and [Add support for add-in commands for Outlook Mobile](add-mobile-support.md).</span></span>

## <a name="implement-adding-online-meeting-details"></a><span data-ttu-id="50193-119">Реализация добавления сведений о собрании по сети</span><span class="sxs-lookup"><span data-stu-id="50193-119">Implement adding online meeting details</span></span>

<span data-ttu-id="50193-120">В этом разделе описывается, как сценарий надстройки может обновить собрание пользователя, включив сведения о собрании по сети.</span><span class="sxs-lookup"><span data-stu-id="50193-120">In this section, learn how your add-in script can update a user's meeting to include online meeting details.</span></span>

1. <span data-ttu-id="50193-121">В проекте быстрого запуска откройте **commands.jsфайл./СРК/коммандс/** в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="50193-121">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="50193-122">Замените весь контент файла **commands.js** на следующий код JavaScript.</span><span class="sxs-lookup"><span data-stu-id="50193-122">Replace the entire content of the **commands.js** file with the following JavaScript.</span></span>

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

## <a name="testing-and-validation"></a><span data-ttu-id="50193-123">Тестирование и проверка</span><span class="sxs-lookup"><span data-stu-id="50193-123">Testing and validation</span></span>

<span data-ttu-id="50193-124">Следуйте обычным рекомендациям по [тестированию и проверке надстройки](testing-and-tips.md).</span><span class="sxs-lookup"><span data-stu-id="50193-124">Follow the usual guidance to [test and validate your add-in](testing-and-tips.md).</span></span> <span data-ttu-id="50193-125">После [загрузки неопубликованных приложений](sideload-outlook-add-ins-for-testing.md) в Outlook в Интернете, Windows или Mac перезапустите Outlook на мобильном устройстве с Android.</span><span class="sxs-lookup"><span data-stu-id="50193-125">After [sideloading](sideload-outlook-add-ins-for-testing.md) in Outlook on the web, Windows, or Mac, restart Outlook on your Android mobile device.</span></span> <span data-ttu-id="50193-126">(Android это единственный поддерживаемый клиент для сейчас.) Затем на новом экране собрания убедитесь, что переключатель Microsoft Teams или Skype заменяется вашим собственным.</span><span class="sxs-lookup"><span data-stu-id="50193-126">(Android is the only supported client for now.) Then, on a new meeting screen, verify that the Microsoft Teams or Skype toggle is replaced with your own.</span></span>

### <a name="create-meeting-ui"></a><span data-ttu-id="50193-127">Создание пользовательского интерфейса собрания</span><span class="sxs-lookup"><span data-stu-id="50193-127">Create meeting UI</span></span>

<span data-ttu-id="50193-128">Как организатор собрания, при создании собрания должны отображаться растры, аналогичные следующим трем изображениям.</span><span class="sxs-lookup"><span data-stu-id="50193-128">As a meeting organizer, you should see screens similar to the following three images when you create a meeting.</span></span>

<span data-ttu-id="50193-129">снимок экрана: " [ ![ Создание экрана собрания" на странице "переключить](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) снимок экрана" на [ ![ странице "Создание собрания" на странице "Создание собрания" на](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) странице "Создание собрания" на странице " [ ![ Создание собрания](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox) " на панели Android/contoso</span><span class="sxs-lookup"><span data-stu-id="50193-129">[![screenshot of create meeting screen on Android - Contoso toggle off](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [![screenshot of create meeting screen on Android - loading Contoso toggle](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [![screenshot of create meeting screen on Android - Contoso toggle on](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)</span></span>

### <a name="join-meeting-ui"></a><span data-ttu-id="50193-130">Пользовательский интерфейс присоединения к собранию</span><span class="sxs-lookup"><span data-stu-id="50193-130">Join meeting UI</span></span>

<span data-ttu-id="50193-131">При просмотре собрания в качестве участника собрания должен отображаться экран, аналогичный следующему изображению.</span><span class="sxs-lookup"><span data-stu-id="50193-131">As a meeting attendee, you should see a screen similar to the following image when you view the meeting.</span></span>

<span data-ttu-id="50193-132">[![снимок экрана с экраном "присоединение к собранию" на Android](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span><span class="sxs-lookup"><span data-stu-id="50193-132">[![screenshot of join meeting screen on Android](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="50193-133">Если вы не видите ссылку **присоединиться** , возможно, на наших серверах не зарегистрирован шаблон собрания в Интернете для вашей службы.</span><span class="sxs-lookup"><span data-stu-id="50193-133">If you don't see the **Join** link, it may be that the online-meeting template for your service is not registered on our servers.</span></span> <span data-ttu-id="50193-134">Подробные сведения можно найти в разделе [Register The Online Template (шаблон собрания)](#register-your-online-meeting-template) .</span><span class="sxs-lookup"><span data-stu-id="50193-134">See the [Register your online-meeting template](#register-your-online-meeting-template) section for details.</span></span>

## <a name="register-your-online-meeting-template"></a><span data-ttu-id="50193-135">Регистрация шаблона собрания в Интернете</span><span class="sxs-lookup"><span data-stu-id="50193-135">Register your online-meeting template</span></span>

<span data-ttu-id="50193-136">Если вы хотите зарегистрировать шаблон собрания в Интернете для своей службы, вы можете создать ошибку GitHub с подробными сведениями.</span><span class="sxs-lookup"><span data-stu-id="50193-136">If you would like to register the online-meeting template for your service, you can create a GitHub issue with the details.</span></span> <span data-ttu-id="50193-137">После этого мы свяжемся с вами, чтобы координировать временную шкалу регистрации.</span><span class="sxs-lookup"><span data-stu-id="50193-137">After that, we'll contact you to coordinate registration timeline.</span></span>

1. <span data-ttu-id="50193-138">Перейдите к разделу **Отзывы** в конце этой статьи.</span><span class="sxs-lookup"><span data-stu-id="50193-138">Go to the **Feedback** section at the end of this article.</span></span>
1. <span data-ttu-id="50193-139">Нажмите ссылку на **эту страницу** .</span><span class="sxs-lookup"><span data-stu-id="50193-139">Press the **This page** link.</span></span>
1. <span data-ttu-id="50193-140">Задайте **название** новой неисправности "зарегистрировать шаблон собрания в сети для My-Service", заменив его на `my-service` имя службы.</span><span class="sxs-lookup"><span data-stu-id="50193-140">Set the **Title** of the new issue to "Register the online-meeting template for my-service", replacing `my-service` with your service name.</span></span>
1. <span data-ttu-id="50193-141">В тексте вопроса замените строку "[Введите здесь обратную связь]" на строку, указанную в `newBody` переменной или аналогичной переменной в разделе [Реализация Добавление сведений о собрании по сети](#implement-adding-online-meeting-details) ранее в этой статье.</span><span class="sxs-lookup"><span data-stu-id="50193-141">In the issue body, replace the string "[Enter feedback here]" with the string you set in the `newBody` or similar variable from the [Implement adding online meeting details](#implement-adding-online-meeting-details) section earlier in this article.</span></span>
1. <span data-ttu-id="50193-142">Нажмите кнопку **Добавить новую ошибку**.</span><span class="sxs-lookup"><span data-stu-id="50193-142">Click **Submit new issue**.</span></span>

![снимок экрана с новым экраном о проблемах GitHub с образцом контента contoso](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a><span data-ttu-id="50193-144">Доступные API</span><span class="sxs-lookup"><span data-stu-id="50193-144">Available APIs</span></span>

<span data-ttu-id="50193-145">Для этой функции доступны следующие API.</span><span class="sxs-lookup"><span data-stu-id="50193-145">The following APIs are available for this feature.</span></span>

- <span data-ttu-id="50193-146">API организатора встречи</span><span class="sxs-lookup"><span data-stu-id="50193-146">Appointment Organizer APIs</span></span>
  - <span data-ttu-id="50193-147">[Office. Context. Mailbox. Item. subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject) ([subject](/javascript/api/outlook/office.subject?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="50193-147">[Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="50193-148">[Office. Context. Mailbox. Item. Start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start) ([время](/javascript/api/outlook/office.time?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="50193-148">[Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="50193-149">[Office. Context. Mailbox. Item. end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end) ([время](/javascript/api/outlook/office.time?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="50193-149">[Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="50193-150">[Office. Context. Mailbox. Item. Location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="50193-150">[Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="50193-151">[Office. Context. Mailbox. Item. optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="50193-151">[Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="50193-152">[Office. Context. Mailbox. Item. requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="50193-152">[Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="50193-153">[Office. Context. Mailbox. Item. Body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body) ([Body. onasync](/javascript/api/outlook/office.body?view=outlook-js-preview#getasync-coerciontype--options--callback-), [Body. setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#setasync-data--options--callback-))</span><span class="sxs-lookup"><span data-stu-id="50193-153">[Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#getasync-coerciontype--options--callback-), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#setasync-data--options--callback-))</span></span>
  - <span data-ttu-id="50193-154">[Office. Context. Mailbox. Item. loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="50193-154">[Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview))</span></span>
  - <span data-ttu-id="50193-155">[Office. Context. roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview#roamingsettings-roamingsettings) ([roamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview))</span><span class="sxs-lookup"><span data-stu-id="50193-155">[Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview))</span></span>
- <span data-ttu-id="50193-156">Обработка процесса проверки подлинности</span><span class="sxs-lookup"><span data-stu-id="50193-156">Handle auth flow</span></span>
  - [<span data-ttu-id="50193-157">API диалоговых окон</span><span class="sxs-lookup"><span data-stu-id="50193-157">Dialog APIs</span></span>](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a><span data-ttu-id="50193-158">Наложен</span><span class="sxs-lookup"><span data-stu-id="50193-158">Restrictions</span></span>

<span data-ttu-id="50193-159">Применяются некоторые ограничения.</span><span class="sxs-lookup"><span data-stu-id="50193-159">Several restrictions apply.</span></span>

- <span data-ttu-id="50193-160">Применяется только к поставщикам услуг для собраний по сети.</span><span class="sxs-lookup"><span data-stu-id="50193-160">Applicable only to online-meeting service providers.</span></span>
- <span data-ttu-id="50193-161">В настоящее время Android является единственным поддерживаемым клиентом.</span><span class="sxs-lookup"><span data-stu-id="50193-161">At present, Android is the only supported client.</span></span> <span data-ttu-id="50193-162">Поддержка iOS скоро будет доступна.</span><span class="sxs-lookup"><span data-stu-id="50193-162">Support on iOS is coming soon.</span></span>
- <span data-ttu-id="50193-163">Только надстройки, установленные администратором, будут отображаться на экране создания собрания, заменив параметры группы по умолчанию или Skype.</span><span class="sxs-lookup"><span data-stu-id="50193-163">Only admin-installed add-ins will appear on the meeting compose screen, replacing the default Teams or Skype option.</span></span> <span data-ttu-id="50193-164">Надстройки, установленные пользователем, не будут активированы.</span><span class="sxs-lookup"><span data-stu-id="50193-164">User-installed add-ins won't activate.</span></span>
- <span data-ttu-id="50193-165">Значок надстройки должен быть в оттенках серого с использованием шестнадцатеричного кода `#919191` или его эквивалента в [других цветовых форматах](https://convertingcolors.com/hex-color-919191.html).</span><span class="sxs-lookup"><span data-stu-id="50193-165">The add-in icon should be in grayscale using hex code `#919191` or its equivalent in [other color formats](https://convertingcolors.com/hex-color-919191.html).</span></span>
- <span data-ttu-id="50193-166">В режиме организатора встречи (создания) поддерживается только одна команда без пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="50193-166">Only one UI-less command is supported in Appointment Organizer (compose) mode.</span></span>

## <a name="see-also"></a><span data-ttu-id="50193-167">См. также</span><span class="sxs-lookup"><span data-stu-id="50193-167">See also</span></span>

- [<span data-ttu-id="50193-168">Надстройки для Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="50193-168">Add-ins for Outlook Mobile</span></span>](outlook-mobile-addins.md)
- [<span data-ttu-id="50193-169">Добавление поддержки команд надстроек для Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="50193-169">Add support for add-in commands for Outlook Mobile</span></span>](add-mobile-support.md)
