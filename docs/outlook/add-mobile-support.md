---
title: Добавление поддержки мобильных устройств в надстройку Outlook
description: Чтобы добавить поддержку Outlook Mobile, необходимо обновить манифест надстройки и, возможно, изменить код для мобильных сценариев.
ms.date: 04/10/2020
localization_priority: Normal
ms.openlocfilehash: f653f43228c7667bc6848d4f0a6d2e9fd1768964
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349009"
---
# <a name="add-support-for-add-in-commands-for-outlook-mobile"></a><span data-ttu-id="b616f-103">Добавление поддержки команд надстроек для Outlook Mobile</span><span class="sxs-lookup"><span data-stu-id="b616f-103">Add support for add-in commands for Outlook Mobile</span></span>

<span data-ttu-id="b616f-104">Использование команд надстройки в Outlook Mobile позволяет пользователям получать доступ к [](#code-considerations)той же функции (с некоторыми ограничениями), что и в Outlook в Интернете, Windows и Mac.</span><span class="sxs-lookup"><span data-stu-id="b616f-104">Using add-in commands in Outlook Mobile allows your users to access the same functionality (with some [limitations](#code-considerations)) that they already have in Outlook on the web, Windows, and Mac.</span></span> <span data-ttu-id="b616f-105">Чтобы добавить поддержку Outlook Mobile, необходимо обновить манифест надстройки и, возможно, изменить код для мобильных сценариев.</span><span class="sxs-lookup"><span data-stu-id="b616f-105">Adding support for Outlook Mobile requires updating the add-in manifest and possibly changing your code for mobile scenarios.</span></span>

## <a name="updating-the-manifest"></a><span data-ttu-id="b616f-106">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="b616f-106">Updating the manifest</span></span>

<span data-ttu-id="b616f-p102">Чтобы включить команды надстроек в Outlook Mobile, необходимо сначала определить их в манифесте надстройки. В схеме [VersionOverrides](../reference/manifest/versionoverrides.md) версии 1.1 определен новый форм-фактор для мобильных устройств — [MobileFormFactor](../reference/manifest/mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="b616f-p102">The first step to enabling add-in commands in Outlook Mobile is to define them in the add-in manifest. The [VersionOverrides](../reference/manifest/versionoverrides.md) v1.1 schema defines a new form factor for mobile, [MobileFormFactor](../reference/manifest/mobileformfactor.md).</span></span>

<span data-ttu-id="b616f-p103">Этот элемент содержит все данные для загрузки надстройки в мобильных клиентах. Это позволяет определять совершенно другие элементы пользовательского интерфейса и файлы JavaScript для мобильной версии.</span><span class="sxs-lookup"><span data-stu-id="b616f-p103">This element contains all of the information for loading the add-in in mobile clients. This enables you to define completely different UI elements and JavaScript files for the mobile experience.</span></span>

<span data-ttu-id="b616f-111">В следующем примере показана одна кнопка области задач в `MobileFormFactor` элементе.</span><span class="sxs-lookup"><span data-stu-id="b616f-111">The following example shows a single task pane button in a `MobileFormFactor` element.</span></span>

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
  ...
  <MobileFormFactor>
    <FunctionFile resid="residUILessFunctionFileUrl" />
    <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
      <Group id="mobileMsgRead">
        <Label resid="groupLabel" />
        <Control xsi:type="MobileButton" id="TaskPaneBtn">
          <Label resid="residTaskPaneButtonName" />
          <Icon xsi:type="bt:MobileIconList">
            <bt:Image size="25" scale="1" resid="tp0icon" />
            <bt:Image size="25" scale="2" resid="tp0icon" />
            <bt:Image size="25" scale="3" resid="tp0icon" />

            <bt:Image size="32" scale="1" resid="tp0icon" />
            <bt:Image size="32" scale="2" resid="tp0icon" />
            <bt:Image size="32" scale="3" resid="tp0icon" />

            <bt:Image size="48" scale="1" resid="tp0icon" />
            <bt:Image size="48" scale="2" resid="tp0icon" />
            <bt:Image size="48" scale="3" resid="tp0icon" />
          </Icon>
          <Action xsi:type="ShowTaskpane">
            <SourceLocation resid="residTaskpaneUrl" />
          </Action>
        </Control>
      </Group>
    </ExtensionPoint>
  </MobileFormFactor>
  ...
</VersionOverrides>
```

<span data-ttu-id="b616f-112">Она во многом подобна элементам, которые отображаются в элементе [DesktopFormFactor](../reference/manifest/desktopformfactor.md), но имеет некоторые существенные отличия.</span><span class="sxs-lookup"><span data-stu-id="b616f-112">This is very similar to the elements that appear in a [DesktopFormFactor](../reference/manifest/desktopformfactor.md) element, with some notable differences.</span></span>

- <span data-ttu-id="b616f-113">Элемент [OfficeTab](../reference/manifest/officetab.md) не используется.</span><span class="sxs-lookup"><span data-stu-id="b616f-113">The [OfficeTab](../reference/manifest/officetab.md) element is not used.</span></span>
- <span data-ttu-id="b616f-p104">У элемента [ExtensionPoint](../reference/manifest/extensionpoint.md) должен быть только один дочерний элемент. Если надстройка добавляет только одну кнопку, это должен быть дочерний элемент [Control](../reference/manifest/control.md). Если же надстройка добавляет несколько кнопок, это должен быть дочерний элемент [Group](../reference/manifest/group.md), содержащий несколько элементов `Control`.</span><span class="sxs-lookup"><span data-stu-id="b616f-p104">The [ExtensionPoint](../reference/manifest/extensionpoint.md) element must have only one child element. If the add-in only adds one button, the child element should be a [Control](../reference/manifest/control.md) element. If the add-in adds more than one button, the child element should be a [Group](../reference/manifest/group.md) element that contains multiple `Control` elements.</span></span>
- <span data-ttu-id="b616f-117">Для элемента `Menu` нет аналога типа `Control`.</span><span class="sxs-lookup"><span data-stu-id="b616f-117">There is no `Menu` type equivalent for the `Control` element.</span></span>
- <span data-ttu-id="b616f-118">Элемент [Supertip](../reference/manifest/supertip.md) не используется.</span><span class="sxs-lookup"><span data-stu-id="b616f-118">The [Supertip](../reference/manifest/supertip.md) element is not used.</span></span>
- <span data-ttu-id="b616f-p105">Требуются значки других размеров. Мобильные надстройки должны поддерживать как минимум значки размерами 25x25, 32x32 и 48x48 пикселей.</span><span class="sxs-lookup"><span data-stu-id="b616f-p105">The required icon sizes are different. Mobile add-ins minimally must support 25x25, 32x32 and 48x48 pixel icons.</span></span>

## <a name="code-considerations"></a><span data-ttu-id="b616f-121">Особенности кода</span><span class="sxs-lookup"><span data-stu-id="b616f-121">Code considerations</span></span>

<span data-ttu-id="b616f-122">При разработке надстроек для мобильных устройств возникают некоторые дополнительные особенности.</span><span class="sxs-lookup"><span data-stu-id="b616f-122">Designing an add-in for mobile introduces some additional considerations.</span></span>

### <a name="use-rest-instead-of-exchange-web-services"></a><span data-ttu-id="b616f-123">Использование REST вместо веб-служб Exchange</span><span class="sxs-lookup"><span data-stu-id="b616f-123">Use REST instead of Exchange Web Services</span></span>

<span data-ttu-id="b616f-p106">Метод [Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) не поддерживается в Outlook Mobile. По мере возможности надстройки должны отдавать предпочтение данным из API Office.js. Если надстройкам требуются сведения, которые не предоставляет API Office.js, то для доступа к почтовому ящику пользователя следует использовать [интерфейсы REST API Outlook](/outlook/rest/).</span><span class="sxs-lookup"><span data-stu-id="b616f-p106">The [Office.context.mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method is not supported in Outlook Mobile. Add-ins should prefer to get information from the Office.js API when possible. If add-ins require information not exposed by the Office.js API, then they should use the [Outlook REST APIs](/outlook/rest/) to access the user's mailbox.</span></span>

<span data-ttu-id="b616f-127">Набор требований к почтовым ящикам 1.5 представил новую версию [Office.context.mailbox.getCallbackTokenAsync,](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) которая может запрашивать маркер доступа, совместимый с API REST, и новое [свойство Office.context.mailbox.restUrl,](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) которое можно использовать для поиска конечной точки API REST для пользователя.</span><span class="sxs-lookup"><span data-stu-id="b616f-127">Mailbox requirement set 1.5 introduced a new version of [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) that can request an access token compatible with the REST APIs, and a new [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) property that can be used to find the REST API endpoint for the user.</span></span>

### <a name="pinch-zoom"></a><span data-ttu-id="b616f-128">Масштабирование жестами</span><span class="sxs-lookup"><span data-stu-id="b616f-128">Pinch zoom</span></span>

<span data-ttu-id="b616f-p107">По умолчанию пользователи могут приближать области задач с помощью жеста масштабирования. Если в вашем случае это неуместно, отключите масштабирование жестами в коде HTML.</span><span class="sxs-lookup"><span data-stu-id="b616f-p107">By default users can use the "pinch zoom" gesture to zoom in on task panes. If this does not make sense for your scenario, be sure to disable pinch zoom in your HTML.</span></span>

### <a name="close-task-panes"></a><span data-ttu-id="b616f-131">Закрытие области задач</span><span class="sxs-lookup"><span data-stu-id="b616f-131">Close task panes</span></span>

<span data-ttu-id="b616f-p108">В Outlook Mobile области задач занимают весь экран, поэтому для возврата к сообщению их необходимо закрывать. Рекомендуем использовать метод [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--), чтобы закрыть область задач по завершении сценария.</span><span class="sxs-lookup"><span data-stu-id="b616f-p108">In Outlook Mobile, task panes take up the entire screen and by default require the user to close them to return to the message. Consider using the [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--) method to close the task pane when your scenario is complete.</span></span>

### <a name="compose-mode-and-appointments"></a><span data-ttu-id="b616f-134">Режим создания и встречи</span><span class="sxs-lookup"><span data-stu-id="b616f-134">Compose mode and appointments</span></span>

<span data-ttu-id="b616f-135">В настоящее время надстройки Outlook Mobile поддерживают активацию только при просмотре сообщений.</span><span class="sxs-lookup"><span data-stu-id="b616f-135">Currently add-ins in Outlook Mobile only support activation when reading messages.</span></span> <span data-ttu-id="b616f-136">Надстройки не активируются при создании сообщений, а также при просмотре и создании встреч.</span><span class="sxs-lookup"><span data-stu-id="b616f-136">Add-ins are not activated when composing messages or when viewing or composing appointments.</span></span> <span data-ttu-id="b616f-137">Однако интегрированные надстройки поставщика собраний в Интернете можно активировать в режиме Организатор встречи.</span><span class="sxs-lookup"><span data-stu-id="b616f-137">However, online meeting provider integrated add-ins can be activated in Appointment Organizer mode.</span></span> <span data-ttu-id="b616f-138">Дополнительные [статьи об этом исключении см. в статье Создание](online-meeting.md) надстройки Outlook для поставщика веб-собраний.</span><span class="sxs-lookup"><span data-stu-id="b616f-138">See the [Create an Outlook mobile add-in for an online-meeting provider](online-meeting.md) article for more about this exception.</span></span>

### <a name="unsupported-apis"></a><span data-ttu-id="b616f-139">Неподдерживаемые интерфейсы API</span><span class="sxs-lookup"><span data-stu-id="b616f-139">Unsupported APIs</span></span>

<span data-ttu-id="b616f-140">API, введенные в наборе требований 1.6 или более поздней, не поддерживаются Outlook Mobile.</span><span class="sxs-lookup"><span data-stu-id="b616f-140">APIs introduced in requirement set 1.6 or later are not supported by Outlook Mobile.</span></span> <span data-ttu-id="b616f-141">Следующие API из предыдущих наборов требований также не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="b616f-141">The following APIs from earlier requirement sets are also not supported.</span></span>

- [<span data-ttu-id="b616f-142">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="b616f-142">Office.context.officeTheme</span></span>](../reference/objectmodel/preview-requirement-set/office.context.md#officetheme-officetheme)
- [<span data-ttu-id="b616f-143">Office.context.mailbox.ewsUrl</span><span class="sxs-lookup"><span data-stu-id="b616f-143">Office.context.mailbox.ewsUrl</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)
- [<span data-ttu-id="b616f-144">Office.context.mailbox.convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="b616f-144">Office.context.mailbox.convertToEwsId</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
- [<span data-ttu-id="b616f-145">Office.context.mailbox.convertToRestId</span><span class="sxs-lookup"><span data-stu-id="b616f-145">Office.context.mailbox.convertToRestId</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
- [<span data-ttu-id="b616f-146">Office.context.mailbox.displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="b616f-146">Office.context.mailbox.displayAppointmentForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
- [<span data-ttu-id="b616f-147">Office.context.mailbox.displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="b616f-147">Office.context.mailbox.displayMessageForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
- [<span data-ttu-id="b616f-148">Office.context.mailbox.displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="b616f-148">Office.context.mailbox.displayNewAppointmentForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
- [<span data-ttu-id="b616f-149">Office.context.mailbox.makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="b616f-149">Office.context.mailbox.makeEwsRequestAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
- [<span data-ttu-id="b616f-150">Office.context.mailbox.item.dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="b616f-150">Office.context.mailbox.item.dateTimeModified</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
- [<span data-ttu-id="b616f-151">Office.context.mailbox.item.displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="b616f-151">Office.context.mailbox.item.displayReplyAllForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [<span data-ttu-id="b616f-152">Office.context.mailbox.item.displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="b616f-152">Office.context.mailbox.item.displayReplyForm</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [<span data-ttu-id="b616f-153">Office.context.mailbox.item.getEntities</span><span class="sxs-lookup"><span data-stu-id="b616f-153">Office.context.mailbox.item.getEntities</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [<span data-ttu-id="b616f-154">Office.context.mailbox.item.getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="b616f-154">Office.context.mailbox.item.getEntitiesByType</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [<span data-ttu-id="b616f-155">Office.context.mailbox.item.getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="b616f-155">Office.context.mailbox.item.getFilteredEntitiesByName</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [<span data-ttu-id="b616f-156">Office.context.mailbox.item.getRegexMatches</span><span class="sxs-lookup"><span data-stu-id="b616f-156">Office.context.mailbox.item.getRegexMatches</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [<span data-ttu-id="b616f-157">Office.context.mailbox.item.getRegexMatchesByName</span><span class="sxs-lookup"><span data-stu-id="b616f-157">Office.context.mailbox.item.getRegexMatchesByName</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

## <a name="see-also"></a><span data-ttu-id="b616f-158">См. также</span><span class="sxs-lookup"><span data-stu-id="b616f-158">See also</span></span>

[<span data-ttu-id="b616f-159">Наборы обязательных элементов, поддерживаемые серверами Exchange и клиентами Outlook</span><span class="sxs-lookup"><span data-stu-id="b616f-159">Requirement sets supported by Exchange servers and Outlook clients</span></span>](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)