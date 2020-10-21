---
title: Реализация добавления в надстройку Outlook с помощью командлета send
description: Узнайте, как реализовать функцию "присоединение к передаче" в надстройке Outlook.
ms.topic: article
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 62234f580f6ff6be418f1c252510f234e297b0c6
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626458"
---
# <a name="implement-append-on-send-in-your-outlook-add-in"></a><span data-ttu-id="3c4a5-103">Реализация добавления в надстройку Outlook с помощью командлета send</span><span class="sxs-lookup"><span data-stu-id="3c4a5-103">Implement append-on-send in your Outlook add-in</span></span>

<span data-ttu-id="3c4a5-104">По завершении этого пошагового руководства у вас будет надстройка Outlook, которая может вставить заявление об отказе при отправке сообщения.</span><span class="sxs-lookup"><span data-stu-id="3c4a5-104">By the end of this walkthrough, you'll have an Outlook add-in that can insert a disclaimer when a message is sent.</span></span>

> [!NOTE]
> <span data-ttu-id="3c4a5-105">Поддержка этой функции появилась в наборе требований 1,9.</span><span class="sxs-lookup"><span data-stu-id="3c4a5-105">Support for this feature was introduced in requirement set 1.9.</span></span> <span data-ttu-id="3c4a5-106">См [клиенты и платформы](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.</span><span class="sxs-lookup"><span data-stu-id="3c4a5-106">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="3c4a5-107">Настройка среды</span><span class="sxs-lookup"><span data-stu-id="3c4a5-107">Set up your environment</span></span>

<span data-ttu-id="3c4a5-108">Завершите работу с [быстрым запуском Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) , который создает проект надстройки с помощью генератора Yeoman для надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="3c4a5-108">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="3c4a5-109">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="3c4a5-109">Configure the manifest</span></span>

<span data-ttu-id="3c4a5-110">Чтобы включить функцию Append-on-Send в надстройке, необходимо включить `AppendOnSend` разрешение в коллекцию [екстендедпермиссионс](../reference/manifest/extendedpermissions.md).</span><span class="sxs-lookup"><span data-stu-id="3c4a5-110">To enable the append-on-send feature in your add-in, you must include the `AppendOnSend` permission in the collection of [ExtendedPermissions](../reference/manifest/extendedpermissions.md).</span></span>

<span data-ttu-id="3c4a5-111">В этом сценарии вместо того, чтобы запускать `action` функцию при нажатии кнопки **выполнить действие** , вы заработаете `appendOnSend` функцию.</span><span class="sxs-lookup"><span data-stu-id="3c4a5-111">For this scenario, instead of running the `action` function on choosing the **Perform an action** button, you'll be running the `appendOnSend` function.</span></span>

1. <span data-ttu-id="3c4a5-112">В редакторе кода откройте Быстрый запуск проекта.</span><span class="sxs-lookup"><span data-stu-id="3c4a5-112">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="3c4a5-113">Откройте файл **manifest.xml** , расположенный в корневом каталоге проекта.</span><span class="sxs-lookup"><span data-stu-id="3c4a5-113">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="3c4a5-114">Выберите весь `<VersionOverrides>` узел (включая открывающие и закрывающие теги) и замените его следующим XML-документом.</span><span class="sxs-lookup"><span data-stu-id="3c4a5-114">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
> <span data-ttu-id="3c4a5-115">Чтобы узнать больше о манифестах для надстроек Outlook, ознакомьтесь с разделом [манифесты надстроек Outlook](manifests.md).</span><span class="sxs-lookup"><span data-stu-id="3c4a5-115">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-append-on-send-handling"></a><span data-ttu-id="3c4a5-116">Реализация обработки при отправке по требованию</span><span class="sxs-lookup"><span data-stu-id="3c4a5-116">Implement append-on-send handling</span></span>

<span data-ttu-id="3c4a5-117">Затем реализуйте Добавление в событие Send.</span><span class="sxs-lookup"><span data-stu-id="3c4a5-117">Next, implement appending on the send event.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3c4a5-118">Если надстройка также реализует [обработку событий при отправке с помощью `ItemSend` ](outlook-on-send-addins.md), вызов `AppendOnSendAsync` в обработчике on – Send возвращает сообщение об ошибке, так как этот сценарий не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="3c4a5-118">If your add-in also implements [on-send event handling using `ItemSend`](outlook-on-send-addins.md), calling `AppendOnSendAsync` in the on-send handler returns an error as this scenario isn't supported.</span></span>

<span data-ttu-id="3c4a5-119">В этом сценарии вы реализуете Добавление заявления об отказе для элемента при отправке пользователя.</span><span class="sxs-lookup"><span data-stu-id="3c4a5-119">For this scenario, you'll implement appending a disclaimer to the item when the user sends.</span></span>

1. <span data-ttu-id="3c4a5-120">В проекте быстрого запуска откройте \*\*commands.jsфайл./СРК/коммандс/ \*\* в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="3c4a5-120">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="3c4a5-121">После `action` функции вставьте следующую функцию JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3c4a5-121">After the `action` function, insert the following JavaScript function.</span></span>

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

1. <span data-ttu-id="3c4a5-122">В конце файла добавьте следующий оператор:</span><span class="sxs-lookup"><span data-stu-id="3c4a5-122">At the end of the file, add the following statement.</span></span>

    ```js
    g.appendDisclaimerOnSend = appendDisclaimerOnSend;
    ```

## <a name="try-it-out"></a><span data-ttu-id="3c4a5-123">Проверка</span><span class="sxs-lookup"><span data-stu-id="3c4a5-123">Try it out</span></span>

1. <span data-ttu-id="3c4a5-124">Выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="3c4a5-124">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="3c4a5-125">При выполнении этой команды локальный веб-сервер запустится, если он еще не запущен.</span><span class="sxs-lookup"><span data-stu-id="3c4a5-125">When you run this command, the local web server will start if it's not already running.</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="3c4a5-126">Следуйте инструкциям в статье [Загрузка неопубликованных надстройки Outlook для тестирования](sideload-outlook-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="3c4a5-126">Follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md).</span></span>

1. <span data-ttu-id="3c4a5-127">Создайте новое сообщение и добавьте себя в строку " **Кому** ".</span><span class="sxs-lookup"><span data-stu-id="3c4a5-127">Create a new message, and add yourself to the **To** line.</span></span>

1. <span data-ttu-id="3c4a5-128">В меню лента или переполнение выберите команду **выполнить действие**.</span><span class="sxs-lookup"><span data-stu-id="3c4a5-128">From the ribbon or overflow menu, choose **Perform an action**.</span></span>

1. <span data-ttu-id="3c4a5-129">Отправьте сообщение, а затем откройте его в папке **"Входящие" или "** **Отправленные** ", чтобы просмотреть добавленное заявление об отказе.</span><span class="sxs-lookup"><span data-stu-id="3c4a5-129">Send the message, then open it from your **Inbox** or **Sent Items** folder to view the appended disclaimer.</span></span>

    ![Снимок экрана с примером сообщения с сообщением об отказе, добавленном при отправке в Outlook в Интернете.](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a><span data-ttu-id="3c4a5-131">См. также</span><span class="sxs-lookup"><span data-stu-id="3c4a5-131">See also</span></span>

[<span data-ttu-id="3c4a5-132">Манифесты надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="3c4a5-132">Outlook add-in manifests</span></span>](manifests.md)
