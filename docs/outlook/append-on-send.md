---
title: Реализация приложения при отправке в надстройки Outlook
description: Узнайте, как реализовать функцию добавления при отправке в надстройки Outlook.
ms.topic: article
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 8b69fbbaef1d0f060f0675fe5c4948a70d935b7a
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234291"
---
# <a name="implement-append-on-send-in-your-outlook-add-in"></a><span data-ttu-id="0598a-103">Реализация приложения при отправке в надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="0598a-103">Implement append-on-send in your Outlook add-in</span></span>

<span data-ttu-id="0598a-104">К концу этого побочного руководство вы получите надстройку Outlook, которая может вставить заявление об отказе при отправлении сообщения.</span><span class="sxs-lookup"><span data-stu-id="0598a-104">By the end of this walkthrough, you'll have an Outlook add-in that can insert a disclaimer when a message is sent.</span></span>

> [!NOTE]
> <span data-ttu-id="0598a-105">Поддержка этой функции была представлена в наборе требований 1.9.</span><span class="sxs-lookup"><span data-stu-id="0598a-105">Support for this feature was introduced in requirement set 1.9.</span></span> <span data-ttu-id="0598a-106">См [клиенты и платформы](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.</span><span class="sxs-lookup"><span data-stu-id="0598a-106">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="0598a-107">Настройка среды</span><span class="sxs-lookup"><span data-stu-id="0598a-107">Set up your environment</span></span>

<span data-ttu-id="0598a-108">Завершите [краткое начало работы с Outlook,](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) в котором создается проект надстройки с помощью генератора Yeoman для надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="0598a-108">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="0598a-109">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="0598a-109">Configure the manifest</span></span>

<span data-ttu-id="0598a-110">Чтобы включить функцию добавления при отправке в надстройке, необходимо включить разрешение в коллекцию `AppendOnSend` [ExtendedPermissions.](../reference/manifest/extendedpermissions.md)</span><span class="sxs-lookup"><span data-stu-id="0598a-110">To enable the append-on-send feature in your add-in, you must include the `AppendOnSend` permission in the collection of [ExtendedPermissions](../reference/manifest/extendedpermissions.md).</span></span>

<span data-ttu-id="0598a-111">В этом сценарии вместо запуска функции при нажатии кнопки действия вы `action` будете запускать эту  `appendOnSend` функцию.</span><span class="sxs-lookup"><span data-stu-id="0598a-111">For this scenario, instead of running the `action` function on choosing the **Perform an action** button, you'll be running the `appendOnSend` function.</span></span>

1. <span data-ttu-id="0598a-112">В редакторе кода откройте проект быстрого запуска.</span><span class="sxs-lookup"><span data-stu-id="0598a-112">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="0598a-113">Откройте файл **manifest.xml,** расположенный в корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="0598a-113">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="0598a-114">Выберите весь узел (включая открытые и закрываемые `<VersionOverrides>` теги) и замените его на следующий XML-</span><span class="sxs-lookup"><span data-stu-id="0598a-114">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
            <DesktopFormFactor>
              <FunctionFile resid="Commands.Url" />
              <ExtensionPoint xsi:type="MessageComposeCommandSurface">
                <OfficeTab id="TabDefault">
                  <Group id="msgComposeGroup">
                    <Label resid="GroupLabel" />
                    <Control xsi:type="Button" id="msgComposeOpenPaneButton">
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
                        <FunctionName>appendDisclaimerOnSend</FunctionName>
                      </Action>
                    </Control>
                  </Group>
                </OfficeTab>
              </ExtensionPoint>

              <!-- Configure AppointmentOrganizerCommandSurface extension point to support
              append on sending a new appointment. -->

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
        <ExtendedPermissions>
          <ExtendedPermission>AppendOnSend</ExtendedPermission>
        </ExtendedPermissions>
      </VersionOverrides>
    </VersionOverrides>
    ```

> [!TIP]
> <span data-ttu-id="0598a-115">Подробнее о манифестах надстройки Outlook см. в манифестах [надстройки Outlook.](manifests.md)</span><span class="sxs-lookup"><span data-stu-id="0598a-115">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-append-on-send-handling"></a><span data-ttu-id="0598a-116">Реализация обработки приложений при отправке</span><span class="sxs-lookup"><span data-stu-id="0598a-116">Implement append-on-send handling</span></span>

<span data-ttu-id="0598a-117">Затем реализуйте приложение для события отправки.</span><span class="sxs-lookup"><span data-stu-id="0598a-117">Next, implement appending on the send event.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0598a-118">Если надстройка [ `ItemSend` ](outlook-on-send-addins.md)также реализует обработку событий при отправке с помощью, вызов в обработчике при отправке возвращает ошибку, так как этот сценарий `AppendOnSendAsync` не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="0598a-118">If your add-in also implements [on-send event handling using `ItemSend`](outlook-on-send-addins.md), calling `AppendOnSendAsync` in the on-send handler returns an error as this scenario isn't supported.</span></span>

<span data-ttu-id="0598a-119">В этом сценарии при отправке пользователем к элементу будет реализовано заявление об отказе.</span><span class="sxs-lookup"><span data-stu-id="0598a-119">For this scenario, you'll implement appending a disclaimer to the item when the user sends.</span></span>

1. <span data-ttu-id="0598a-120">В том же проекте быстрого запуска откройте файл **./src/commands/commands.js** в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="0598a-120">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="0598a-121">После `action` функции вставьте следующую функцию JavaScript.</span><span class="sxs-lookup"><span data-stu-id="0598a-121">After the `action` function, insert the following JavaScript function.</span></span>

    ```js
    function appendDisclaimerOnSend(event) {
      var appendText =
        '<p style = "color:blue"> <i>This and subsequent emails on the same topic are for discussion and information purposes only. Only those matters set out in a fully executed agreement are legally binding. This email may contain confidential information and should not be shared with any third party without the prior written agreement of Contoso. If you are not the intended recipient, take no action and contact the sender immediately.<br><br>Contoso Limited (company number 01624297) is a company registered in England and Wales whose registered office is at Contoso Campus, Thames Valley Park, Reading RG6 1WG</i></p>';  
      /**
        *************************************************************
         Ideal Usage - Call the getBodyType API. Use the coercionType
         it returns as the parameter value below.
        *************************************************************
      */
      Office.context.mailbox.item.body.appendOnSendAsync(
        appendText,
        {
          coercionType: Office.CoercionType.Html
        },
        function(asyncResult) {
          console.log(asyncResult);
        }
      );

      event.completed();
    }
    ```

1. <span data-ttu-id="0598a-122">В конце файла добавьте следующую выписку.</span><span class="sxs-lookup"><span data-stu-id="0598a-122">At the end of the file, add the following statement.</span></span>

    ```js
    g.appendDisclaimerOnSend = appendDisclaimerOnSend;
    ```

## <a name="try-it-out"></a><span data-ttu-id="0598a-123">Проверка</span><span class="sxs-lookup"><span data-stu-id="0598a-123">Try it out</span></span>

1. <span data-ttu-id="0598a-124">Выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="0598a-124">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="0598a-125">При запуске этой команды запустится локальный веб-сервер, если он еще не запущен и ваша надстройка будет загружена.</span><span class="sxs-lookup"><span data-stu-id="0598a-125">When you run this command, the local web server will start if it's not already running and your add-in will be sideloaded.</span></span> 

    ```command&nbsp;line
    npm start
    ```

1. <span data-ttu-id="0598a-126">Создайте новое сообщение и добавьте себя в **строку "To".**</span><span class="sxs-lookup"><span data-stu-id="0598a-126">Create a new message, and add yourself to the **To** line.</span></span>

1. <span data-ttu-id="0598a-127">На ленте или в меню переполнения выберите **"Выполнить действие".**</span><span class="sxs-lookup"><span data-stu-id="0598a-127">From the ribbon or overflow menu, choose **Perform an action**.</span></span>

1. <span data-ttu-id="0598a-128">Отправьте сообщение, а затем  откройте его  из папки "Входящие" или "Отправленные", чтобы просмотреть заявление об отказе.</span><span class="sxs-lookup"><span data-stu-id="0598a-128">Send the message, then open it from your **Inbox** or **Sent Items** folder to view the appended disclaimer.</span></span>

    ![Снимок экрана с примером сообщения с заявлением об отказе при отправке в Outlook в Интернете.](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a><span data-ttu-id="0598a-130">См. также</span><span class="sxs-lookup"><span data-stu-id="0598a-130">See also</span></span>

[<span data-ttu-id="0598a-131">Манифесты надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="0598a-131">Outlook add-in manifests</span></span>](manifests.md)
