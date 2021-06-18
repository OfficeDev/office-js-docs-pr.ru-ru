---
title: Включить общие папки и сценарии общих почтовых ящиков в Outlook надстройке
description: Обсуждается настройка поддержки надстройки для общих папок (ака). делегирования доступа) и общих почтовых ящиков.
ms.date: 06/17/2021
localization_priority: Normal
ms.openlocfilehash: 5d7fb712b8f814184c2a444c32416d35fb1da49c
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007771"
---
# <a name="enable-shared-folders-and-shared-mailbox-scenarios-in-an-outlook-add-in"></a><span data-ttu-id="f10a6-104">Включить общие папки и сценарии общих почтовых ящиков в Outlook надстройке</span><span class="sxs-lookup"><span data-stu-id="f10a6-104">Enable shared folders and shared mailbox scenarios in an Outlook add-in</span></span>

<span data-ttu-id="f10a6-105">В этой статье описывается, как включить общие папки (также известные как доступ к делегатам) и общие почтовые ящики (в настоящее время в предварительном [просмотре)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#shared-mailboxes)сценарии в надстройке Outlook, в том числе разрешения, которые поддерживает API Office JavaScript.</span><span class="sxs-lookup"><span data-stu-id="f10a6-105">This article describes how to enable shared folders (also known as delegate access) and shared mailbox (now in [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#shared-mailboxes)) scenarios in your Outlook add-in, including which permissions the Office JavaScript API supports.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f10a6-106">Поддержка этой функции была представлена в [наборе требований 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="f10a6-106">Support for this feature was introduced in [requirement set 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span> <span data-ttu-id="f10a6-107">См [клиенты и платформы](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.</span><span class="sxs-lookup"><span data-stu-id="f10a6-107">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-setups"></a><span data-ttu-id="f10a6-108">Поддерживаемые установки</span><span class="sxs-lookup"><span data-stu-id="f10a6-108">Supported setups</span></span>

<span data-ttu-id="f10a6-109">В следующих разделах описываются поддерживаемые конфигурации для общих почтовых ящиков (теперь в предварительном просмотре) и общих папок.</span><span class="sxs-lookup"><span data-stu-id="f10a6-109">The following sections describe supported configurations for shared mailboxes (now in preview) and shared folders.</span></span> <span data-ttu-id="f10a6-110">API-функции могут работать не так, как ожидалось в других конфигурациях.</span><span class="sxs-lookup"><span data-stu-id="f10a6-110">The feature APIs may not work as expected in other configurations.</span></span> <span data-ttu-id="f10a6-111">Выберите платформу, на которой необходимо научиться настраивать.</span><span class="sxs-lookup"><span data-stu-id="f10a6-111">Select the platform you'd like to learn how to configure.</span></span>

### <a name="windows"></a>[<span data-ttu-id="f10a6-112">Windows</span><span class="sxs-lookup"><span data-stu-id="f10a6-112">Windows</span></span>](#tab/windows)

#### <a name="shared-folders"></a><span data-ttu-id="f10a6-113">Общие папки</span><span class="sxs-lookup"><span data-stu-id="f10a6-113">Shared folders</span></span>

<span data-ttu-id="f10a6-114">Сначала владелец почтового ящика [должен предоставить доступ к делегату.](https://support.microsoft.com/office/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)</span><span class="sxs-lookup"><span data-stu-id="f10a6-114">The mailbox owner must first [provide access to a delegate](https://support.microsoft.com/office/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926).</span></span> <span data-ttu-id="f10a6-115">Затем делегат должен следовать инструкциям, изложенным в разделе "Добавление почтового ящика другого человека в свой профиль" статьи Управление почтовыми и календарями другого [пользователя.](https://support.microsoft.com/office/manage-another-person-s-mail-and-calendar-items-afb79d6b-2967-43b9-a944-a6b953190af5)</span><span class="sxs-lookup"><span data-stu-id="f10a6-115">The delegate must then follow the instructions outlined in the "Add another person's mailbox to your profile" section of the article [Manage another person's mail and calendar items](https://support.microsoft.com/office/manage-another-person-s-mail-and-calendar-items-afb79d6b-2967-43b9-a944-a6b953190af5).</span></span>

#### <a name="shared-mailboxes-preview"></a><span data-ttu-id="f10a6-116">Общие почтовые ящики (предварительный просмотр)</span><span class="sxs-lookup"><span data-stu-id="f10a6-116">Shared mailboxes (preview)</span></span>

<span data-ttu-id="f10a6-117">Exchange серверов администраторы могут создавать и управлять общими почтовыми ящиками для наборов пользователей для доступа.</span><span class="sxs-lookup"><span data-stu-id="f10a6-117">Exchange server admins can create and manage shared mailboxes for sets of users to access.</span></span> <span data-ttu-id="f10a6-118">В настоящее [время Exchange Online](/exchange/collaboration-exo/shared-mailboxes) является единственной поддерживаемой серверной версией для этой функции.</span><span class="sxs-lookup"><span data-stu-id="f10a6-118">At present, [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) is the only supported server version for this feature.</span></span>

<span data-ttu-id="f10a6-119">Функция Exchange Server, известная как "автомаппирование", по умолчанию [](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook) включается, что означает, что после закрытия и открытия Outlook Outlook общего почтового ящика должен автоматически отображаться общий почтовый ящик.</span><span class="sxs-lookup"><span data-stu-id="f10a6-119">An Exchange Server feature known as "automapping" is on by default which means that subsequently the [shared mailbox should automatically appear](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook) in a user's Outlook app after Outlook has been closed and reopened.</span></span> <span data-ttu-id="f10a6-120">Однако если администратор отключил автомаппирование, пользователь должен следовать инструкциям, описанным в разделе "Добавление общего почтового ящика в Outlook" статьи Open и использовать общий почтовый ящик в [Outlook](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-d94a8e9e-21f1-4240-808b-de9c9c088afd).</span><span class="sxs-lookup"><span data-stu-id="f10a6-120">However, if an admin turned off automapping, the user must follow the manual steps outlined in the "Add a shared mailbox to Outlook" section of the article [Open and use a shared mailbox in Outlook](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-d94a8e9e-21f1-4240-808b-de9c9c088afd).</span></span>

> [!WARNING]
> <span data-ttu-id="f10a6-121">Не **входящие** в общий почтовый ящик с паролем.</span><span class="sxs-lookup"><span data-stu-id="f10a6-121">Do **NOT** sign into the shared mailbox with a password.</span></span> <span data-ttu-id="f10a6-122">В этом случае API-функции не будут работать.</span><span class="sxs-lookup"><span data-stu-id="f10a6-122">The feature APIs won't work in that case.</span></span>

### <a name="web-browser---modern-outlook"></a>[<span data-ttu-id="f10a6-123">Веб-браузер — современная версия Outlook</span><span class="sxs-lookup"><span data-stu-id="f10a6-123">Web browser - modern Outlook</span></span>](#tab/modern)

#### <a name="shared-folders"></a><span data-ttu-id="f10a6-124">Общие папки</span><span class="sxs-lookup"><span data-stu-id="f10a6-124">Shared folders</span></span>

<span data-ttu-id="f10a6-125">Сначала владелец почтового ящика должен предоставить доступ к [делегату,](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) обновив разрешения папок почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="f10a6-125">The mailbox owner must first [provide access to a delegate](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) by updating the mailbox folder permissions.</span></span> <span data-ttu-id="f10a6-126">Затем делегат должен следовать инструкциям, изложенным в разделе "Добавление почтового ящика другого человека в список папки в Outlook Web App" раздела статьи Доступ к почтовому ящику другого [человека.](https://support.microsoft.com/office/access-another-person-s-mailbox-a909ad30-e413-40b5-a487-0ea70b763081)</span><span class="sxs-lookup"><span data-stu-id="f10a6-126">The delegate must then follow the instructions outlined in the "Add another person’s mailbox to your folder list in Outlook Web App" section of the article [Access another person's mailbox](https://support.microsoft.com/office/access-another-person-s-mailbox-a909ad30-e413-40b5-a487-0ea70b763081).</span></span>

#### <a name="shared-mailboxes-preview"></a><span data-ttu-id="f10a6-127">Общие почтовые ящики (предварительный просмотр)</span><span class="sxs-lookup"><span data-stu-id="f10a6-127">Shared mailboxes (preview)</span></span>

<span data-ttu-id="f10a6-128">Exchange серверов администраторы могут создавать и управлять общими почтовыми ящиками для наборов пользователей для доступа.</span><span class="sxs-lookup"><span data-stu-id="f10a6-128">Exchange server admins can create and manage shared mailboxes for sets of users to access.</span></span> <span data-ttu-id="f10a6-129">В настоящее [время Exchange Online](/exchange/collaboration-exo/shared-mailboxes) является единственной поддерживаемой серверной версией для этой функции.</span><span class="sxs-lookup"><span data-stu-id="f10a6-129">At present, [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) is the only supported server version for this feature.</span></span>

<span data-ttu-id="f10a6-130">После получения доступа общий пользователь почтового ящика должен следовать шагам, описанным в разделе "Добавьте общий почтовый ящик, чтобы он отображался в основном почтовом ящике" в статье Open и использовать общий почтовый ящик в [Outlook в Интернете](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-on-the-web-98b5a90d-4e38-415d-a030-f09a4cd28207).</span><span class="sxs-lookup"><span data-stu-id="f10a6-130">After receiving access, a shared mailbox user must follow the steps outlined in the "Add the shared mailbox so it displays under your primary mailbox" section of the article [Open and use a shared mailbox in Outlook on the web](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-on-the-web-98b5a90d-4e38-415d-a030-f09a4cd28207).</span></span>

> [!WARNING]
> <span data-ttu-id="f10a6-131">Не **используйте** другие параметры, такие как "Откройте другой почтовый ящик".</span><span class="sxs-lookup"><span data-stu-id="f10a6-131">Do **NOT** use other options like "Open another mailbox".</span></span> <span data-ttu-id="f10a6-132">API-функции могут работать неправильно.</span><span class="sxs-lookup"><span data-stu-id="f10a6-132">The feature APIs may not work properly then.</span></span>

---

<span data-ttu-id="f10a6-133">Дополнительные сведения о том, где надстройки делают и [](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) не активируются в целом, обратитесь к пунктам почтовых ящиков, доступным в разделе надстройки на странице Outlook обзор надстройки.</span><span class="sxs-lookup"><span data-stu-id="f10a6-133">To learn more about where add-ins do and do not activate in general, refer to the [Mailbox items available to add-ins](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) section of the Outlook add-ins overview page.</span></span>

## <a name="supported-permissions"></a><span data-ttu-id="f10a6-134">Поддерживаемые разрешения</span><span class="sxs-lookup"><span data-stu-id="f10a6-134">Supported permissions</span></span>

<span data-ttu-id="f10a6-135">В следующей таблице описываются разрешения, Office API JavaScript для делегатов и общих пользователей почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="f10a6-135">The following table describes the permissions that the Office JavaScript API supports for delegates and shared mailbox users.</span></span>

|<span data-ttu-id="f10a6-136">Разрешение</span><span class="sxs-lookup"><span data-stu-id="f10a6-136">Permission</span></span>|<span data-ttu-id="f10a6-137">Значение</span><span class="sxs-lookup"><span data-stu-id="f10a6-137">Value</span></span>|<span data-ttu-id="f10a6-138">Описание</span><span class="sxs-lookup"><span data-stu-id="f10a6-138">Description</span></span>|
|---|---:|---|
|<span data-ttu-id="f10a6-139">Чтение</span><span class="sxs-lookup"><span data-stu-id="f10a6-139">Read</span></span>|<span data-ttu-id="f10a6-140">1 (000001)</span><span class="sxs-lookup"><span data-stu-id="f10a6-140">1 (000001)</span></span>|<span data-ttu-id="f10a6-141">Может читать элементы.</span><span class="sxs-lookup"><span data-stu-id="f10a6-141">Can read items.</span></span>|
|<span data-ttu-id="f10a6-142">Запись</span><span class="sxs-lookup"><span data-stu-id="f10a6-142">Write</span></span>|<span data-ttu-id="f10a6-143">2 (000010)</span><span class="sxs-lookup"><span data-stu-id="f10a6-143">2 (000010)</span></span>|<span data-ttu-id="f10a6-144">Можно создавать элементы.</span><span class="sxs-lookup"><span data-stu-id="f10a6-144">Can create items.</span></span>|
|<span data-ttu-id="f10a6-145">DeleteOwn</span><span class="sxs-lookup"><span data-stu-id="f10a6-145">DeleteOwn</span></span>|<span data-ttu-id="f10a6-146">4 (000100)</span><span class="sxs-lookup"><span data-stu-id="f10a6-146">4 (000100)</span></span>|<span data-ttu-id="f10a6-147">Можно удалить только созданные элементы.</span><span class="sxs-lookup"><span data-stu-id="f10a6-147">Can delete only the items they created.</span></span>|
|<span data-ttu-id="f10a6-148">DeleteAll</span><span class="sxs-lookup"><span data-stu-id="f10a6-148">DeleteAll</span></span>|<span data-ttu-id="f10a6-149">8 (001000)</span><span class="sxs-lookup"><span data-stu-id="f10a6-149">8 (001000)</span></span>|<span data-ttu-id="f10a6-150">Может удалять любые элементы.</span><span class="sxs-lookup"><span data-stu-id="f10a6-150">Can delete any items.</span></span>|
|<span data-ttu-id="f10a6-151">EditOwn</span><span class="sxs-lookup"><span data-stu-id="f10a6-151">EditOwn</span></span>|<span data-ttu-id="f10a6-152">16 (010000)</span><span class="sxs-lookup"><span data-stu-id="f10a6-152">16 (010000)</span></span>|<span data-ttu-id="f10a6-153">Можно редактировать только созданные элементы.</span><span class="sxs-lookup"><span data-stu-id="f10a6-153">Can edit only the items they created.</span></span>|
|<span data-ttu-id="f10a6-154">EditAll</span><span class="sxs-lookup"><span data-stu-id="f10a6-154">EditAll</span></span>|<span data-ttu-id="f10a6-155">32 (100000)</span><span class="sxs-lookup"><span data-stu-id="f10a6-155">32 (100000)</span></span>|<span data-ttu-id="f10a6-156">Может изменять любые элементы.</span><span class="sxs-lookup"><span data-stu-id="f10a6-156">Can edit any items.</span></span>|

> [!NOTE]
> <span data-ttu-id="f10a6-157">В настоящее время API поддерживает получение существующих разрешений, но не установку разрешений.</span><span class="sxs-lookup"><span data-stu-id="f10a6-157">Currently the API supports getting existing permissions, but not setting permissions.</span></span>

<span data-ttu-id="f10a6-158">Объект [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) реализуется с помощью битмаски для указать разрешения.</span><span class="sxs-lookup"><span data-stu-id="f10a6-158">The [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) object is implemented using a bitmask to indicate the permissions.</span></span> <span data-ttu-id="f10a6-159">Каждая позиция в битмаске представляет определенное разрешение, и если оно заданной, у пользователя `1` есть соответствующее разрешение.</span><span class="sxs-lookup"><span data-stu-id="f10a6-159">Each position in the bitmask represents a particular permission and if it's set to `1` then the user has the respective permission.</span></span> <span data-ttu-id="f10a6-160">Например, если справа находится второй `1` бит, у пользователя есть разрешение **Напишите.**</span><span class="sxs-lookup"><span data-stu-id="f10a6-160">For example, if the second bit from the right is `1`, then the user has **Write** permission.</span></span> <span data-ttu-id="f10a6-161">Пример проверки определенного разрешения в разделе [Выполнение](#perform-an-operation-as-delegate-or-shared-mailbox-user) операции в качестве делегата или общего пользователя почтового ящика см. в этой статье.</span><span class="sxs-lookup"><span data-stu-id="f10a6-161">You can see an example of how to check for a specific permission in the [Perform an operation as delegate or shared mailbox user](#perform-an-operation-as-delegate-or-shared-mailbox-user) section later in this article.</span></span>

## <a name="sync-across-shared-folder-clients"></a><span data-ttu-id="f10a6-162">Синхронизация между общими клиентами папок</span><span class="sxs-lookup"><span data-stu-id="f10a6-162">Sync across shared folder clients</span></span>

<span data-ttu-id="f10a6-163">Обновления делегата в почтовом ящике владельца обычно синхронизируются между почтовыми ящиками немедленно.</span><span class="sxs-lookup"><span data-stu-id="f10a6-163">A delegate's updates to the owner's mailbox are usually synced across mailboxes immediately.</span></span>

<span data-ttu-id="f10a6-164">Однако если операции REST или Exchange Web Services (EWS) использовались для набора расширенного свойства элемента, синхронизация таких изменений может занять несколько часов. Мы рекомендуем вместо этого использовать [объект CustomProperties](/javascript/api/outlook/office.customproperties) и связанные API, чтобы избежать такой задержки.</span><span class="sxs-lookup"><span data-stu-id="f10a6-164">However, if REST or Exchange Web Services (EWS) operations were used to set an extended property on an item, such changes could take a few hours to sync. We recommend you instead use the [CustomProperties](/javascript/api/outlook/office.customproperties) object and related APIs to avoid such a delay.</span></span> <span data-ttu-id="f10a6-165">Дополнительные дополнительные [](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) статьи см. в разделе настраиваемые свойства в статье "Получить и установить метаданные в Outlook надстройки".</span><span class="sxs-lookup"><span data-stu-id="f10a6-165">To learn more, see the [custom properties section](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) of the "Get and set metadata in an Outlook add-in" article.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f10a6-166">В сценарии делегирования нельзя использовать EWS с маркерами, которые в настоящее время office.js API.</span><span class="sxs-lookup"><span data-stu-id="f10a6-166">In a delegate scenario, you can't use EWS with the tokens currently provided by office.js API.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="f10a6-167">Настройка манифеста</span><span class="sxs-lookup"><span data-stu-id="f10a6-167">Configure the manifest</span></span>

<span data-ttu-id="f10a6-168">Чтобы включить общие папки и сценарии общих почтовых ящиков в надстройке, необходимо настроить элемент [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) в манифесте под `true` родительским элементом. `DesktopFormFactor`</span><span class="sxs-lookup"><span data-stu-id="f10a6-168">To enable shared folders and shared mailbox scenarios in your add-in, you must set the [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) element to `true` in the manifest under the parent element `DesktopFormFactor`.</span></span> <span data-ttu-id="f10a6-169">В настоящее время другие форм-факторы не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="f10a6-169">At present, other form factors are not supported.</span></span>

<span data-ttu-id="f10a6-170">Чтобы поддерживать вызовы REST от делегата, установите узел [Разрешений](../reference/manifest/permissions.md) в `ReadWriteMailbox` манифесте.</span><span class="sxs-lookup"><span data-stu-id="f10a6-170">To support REST calls from a delegate, set the [Permissions](../reference/manifest/permissions.md) node in the manifest to `ReadWriteMailbox`.</span></span>

<span data-ttu-id="f10a6-171">В следующем примере показан элемент, установленный `SupportsSharedFolders` `true` в разделе манифеста.</span><span class="sxs-lookup"><span data-stu-id="f10a6-171">The following example shows the `SupportsSharedFolders` element set to `true` in a section of the manifest.</span></span>

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- configure selected extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="perform-an-operation-as-delegate-or-shared-mailbox-user"></a><span data-ttu-id="f10a6-172">Выполните операцию в качестве пользователя делегирования или общего почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f10a6-172">Perform an operation as delegate or shared mailbox user</span></span>

<span data-ttu-id="f10a6-173">Общие свойства элемента можно получить в режиме Compose или Read, позвонив по методу [item.getSharedPropertiesAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)</span><span class="sxs-lookup"><span data-stu-id="f10a6-173">You can get an item's shared properties in Compose or Read mode by calling the [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="f10a6-174">Это возвращает объект [SharedProperties,](/javascript/api/outlook/office.sharedproperties) который в настоящее время предоставляет разрешения пользователя, адрес электронной почты владельца, базовый URL-адрес API REST и целевой почтовый ящик.</span><span class="sxs-lookup"><span data-stu-id="f10a6-174">This returns a [SharedProperties](/javascript/api/outlook/office.sharedproperties) object that currently provides the user's permissions, the owner's email address, the REST API's base URL, and the target mailbox.</span></span>

<span data-ttu-id="f10a6-175">В следующем примере показано, как получить общие свойства сообщения или встречи, проверить, есть ли у делегата или общего пользователя почтового ящика разрешение на запись, и сделать вызов REST. </span><span class="sxs-lookup"><span data-stu-id="f10a6-175">The following example shows how to get the shared properties of a message or appointment, check if the delegate or shared mailbox user has **Write** permission, and make a REST call.</span></span>

```js
function performOperation() {
  Office.context.mailbox.getCallbackTokenAsync({
      isRest: true
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value !== "") {
        Office.context.mailbox.item.getSharedPropertiesAsync({
            // Pass auth token along.
            asyncContext: asyncResult.value
          },
          function (asyncResult1) {
            let sharedProperties = asyncResult1.value;
            let delegatePermissions = sharedProperties.delegatePermissions;

            // Determine if user can do the expected operation.
            // E.g., do they have Write permission?
            if ((delegatePermissions & Office.MailboxEnums.DelegatePermissions.Write) != 0) {
              // Construct REST URL for your operation.
              // Update <version> placeholder with actual Outlook REST API version e.g. "v2.0".
              // Update <operation> placeholder with actual operation.
              let rest_url = sharedProperties.targetRestUrl + "/<version>/users/" + sharedProperties.targetMailbox + "/<operation>";
  
              $.ajax({
                  url: rest_url,
                  dataType: 'json',
                  headers:
                  {
                    "Authorization": "Bearer " + asyncResult1.asyncContext
                  }
                }
              ).done(
                function (response) {
                  console.log("success");
                }
              ).fail(
                function (error) {
                  console.log("error message");
                }
              );
            }
          }
        );
      }
    }
  );
}
```

> [!TIP]
> <span data-ttu-id="f10a6-176">В качестве делегата можно использовать REST для получения содержимого сообщения Outlook, прикрепленного к элементу Outlook [или групповой публикации.](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)</span><span class="sxs-lookup"><span data-stu-id="f10a6-176">As a delegate, you can use REST to [get the content of an Outlook message attached to an Outlook item or group post](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).</span></span>

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a><span data-ttu-id="f10a6-177">Обработка вызовов REST для общих и не общих элементов</span><span class="sxs-lookup"><span data-stu-id="f10a6-177">Handle calling REST on shared and non-shared items</span></span>

<span data-ttu-id="f10a6-178">Если вы хотите вызвать операцию REST для элемента, является ли этот элемент общим, вы можете использовать API, чтобы определить, является ли элемент `getSharedPropertiesAsync` общим.</span><span class="sxs-lookup"><span data-stu-id="f10a6-178">If you want to call a REST operation on an item, whether or not the item is shared, you can use the `getSharedPropertiesAsync` API to determine if the item is shared.</span></span> <span data-ttu-id="f10a6-179">После этого можно создать URL-адрес REST для операции с помощью соответствующего объекта.</span><span class="sxs-lookup"><span data-stu-id="f10a6-179">After that, you can construct the REST URL for the operation using the appropriate object.</span></span>

```js
if (item.getSharedPropertiesAsync) {
  // In Windows, Mac, and the web client, this indicates a shared item so use SharedProperties properties to construct the REST URL.
  // Add-ins don't activate on shared items in mobile so no need to handle.

  // Perform operation for shared item.
} else {
  // In general, this is not a shared item, so construct the REST URL using info from the Call REST APIs article:
  // https://docs.microsoft.com/office/dev/add-ins/outlook/use-rest-api

  // Perform operation for non-shared item.
}
```

## <a name="limitations"></a><span data-ttu-id="f10a6-180">Ограничения</span><span class="sxs-lookup"><span data-stu-id="f10a6-180">Limitations</span></span>

<span data-ttu-id="f10a6-181">В зависимости от сценариев надстройки существует несколько ограничений, которые следует учитывать при обработке общих папок или общих ситуаций почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="f10a6-181">Depending on your add-in's scenarios, there are a few limitations for you to consider when handling shared folder or shared mailbox situations.</span></span>

### <a name="message-compose-mode"></a><span data-ttu-id="f10a6-182">Режим композитации сообщений</span><span class="sxs-lookup"><span data-stu-id="f10a6-182">Message Compose mode</span></span>

<span data-ttu-id="f10a6-183">В режиме композитации сообщений [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_) не поддерживается в Outlook в Интернете или Windows, если не выполнены следующие условия.</span><span class="sxs-lookup"><span data-stu-id="f10a6-183">In Message Compose mode, [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_) is not supported in Outlook on the web or on Windows unless the following conditions are met.</span></span>

<span data-ttu-id="f10a6-184">а.</span><span class="sxs-lookup"><span data-stu-id="f10a6-184">a.</span></span> <span data-ttu-id="f10a6-185">**Делегирование доступа и общих папок**</span><span class="sxs-lookup"><span data-stu-id="f10a6-185">**Delegate access/Shared folders**</span></span>

1. <span data-ttu-id="f10a6-186">Владелец почтового ящика запускает сообщение.</span><span class="sxs-lookup"><span data-stu-id="f10a6-186">The mailbox owner starts a message.</span></span> <span data-ttu-id="f10a6-187">Это может быть новое сообщение, ответ или форвард.</span><span class="sxs-lookup"><span data-stu-id="f10a6-187">This can be a new message, a reply, or a forward.</span></span>
1. <span data-ttu-id="f10a6-188">Затем сообщение сохраняется, а затем перемещается из собственной папки **Drafts** в папку, доступную делегату.</span><span class="sxs-lookup"><span data-stu-id="f10a6-188">They save the message then move it from their own **Drafts** folder to a folder shared with the delegate.</span></span>
1. <span data-ttu-id="f10a6-189">Делегат открывает черновик из общей папки, а затем продолжает сочинять.</span><span class="sxs-lookup"><span data-stu-id="f10a6-189">The delegate opens the draft from the shared folder then continues composing.</span></span>

<span data-ttu-id="f10a6-190">б.</span><span class="sxs-lookup"><span data-stu-id="f10a6-190">b.</span></span> <span data-ttu-id="f10a6-191">**Общий почтовый ящик**</span><span class="sxs-lookup"><span data-stu-id="f10a6-191">**Shared mailbox**</span></span>

1. <span data-ttu-id="f10a6-192">Пользователь общего почтового ящика запускает сообщение.</span><span class="sxs-lookup"><span data-stu-id="f10a6-192">A shared mailbox user starts a message.</span></span> <span data-ttu-id="f10a6-193">Это может быть новое сообщение, ответ или форвард.</span><span class="sxs-lookup"><span data-stu-id="f10a6-193">This can be a new message, a reply, or a forward.</span></span>
1. <span data-ttu-id="f10a6-194">Затем они сэкономят сообщение из собственной папки **Drafts** в папку в общем почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="f10a6-194">They save the message then move it from their own **Drafts** folder to a folder in the shared mailbox.</span></span>
1. <span data-ttu-id="f10a6-195">Другой пользователь общего почтового ящика открывает черновик из общего почтового ящика, а затем продолжает сочинять.</span><span class="sxs-lookup"><span data-stu-id="f10a6-195">Another shared mailbox user opens the draft from the shared mailbox then continues composing.</span></span>

<span data-ttu-id="f10a6-196">Теперь сообщение находится в общем контексте, и надстройки, поддерживаюные эти общие сценарии, могут получать общие свойства элемента.</span><span class="sxs-lookup"><span data-stu-id="f10a6-196">The message is now in a shared context and add-ins that support these shared scenarios can get the item's shared properties.</span></span> <span data-ttu-id="f10a6-197">После отправки сообщения оно обычно находится в папке  отправленных элементов отправители.</span><span class="sxs-lookup"><span data-stu-id="f10a6-197">After the message has been sent, it's usually found in the sender's **Sent Items** folder.</span></span>

### <a name="rest-and-ews"></a><span data-ttu-id="f10a6-198">REST и EWS</span><span class="sxs-lookup"><span data-stu-id="f10a6-198">REST and EWS</span></span>

<span data-ttu-id="f10a6-199">Ваша надстройка может использовать REST, и необходимо установить разрешение надстройки, чтобы включить доступ REST к почтовому ящику владельца или к общему почтовому ящику, как `ReadWriteMailbox` это применимо.</span><span class="sxs-lookup"><span data-stu-id="f10a6-199">Your add-in can use REST and the add-in's permission must be set to `ReadWriteMailbox` to enable REST access to the owner's mailbox or to the shared mailbox as applicable.</span></span> <span data-ttu-id="f10a6-200">EWS не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="f10a6-200">EWS is not supported.</span></span>

## <a name="see-also"></a><span data-ttu-id="f10a6-201">См. также</span><span class="sxs-lookup"><span data-stu-id="f10a6-201">See also</span></span>

- [<span data-ttu-id="f10a6-202">Разрешить другим пользователям управлять почтой и календарем</span><span class="sxs-lookup"><span data-stu-id="f10a6-202">Allow someone else to manage your mail and calendar</span></span>](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [<span data-ttu-id="f10a6-203">Общий доступ к календарю в Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="f10a6-203">Calendar sharing in Microsoft 365</span></span>](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [<span data-ttu-id="f10a6-204">Добавьте общий почтовый ящик в Outlook</span><span class="sxs-lookup"><span data-stu-id="f10a6-204">Add a shared mailbox to Outlook</span></span>](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)
- [<span data-ttu-id="f10a6-205">Как заказать элементы манифеста</span><span class="sxs-lookup"><span data-stu-id="f10a6-205">How to order manifest elements</span></span>](../develop/manifest-element-ordering.md)
- <span data-ttu-id="f10a6-206">[Маска (вычисления)](https://en.wikipedia.org/wiki/Mask_(computing))</span><span class="sxs-lookup"><span data-stu-id="f10a6-206">[Mask (computing)](https://en.wikipedia.org/wiki/Mask_(computing))</span></span>
- [<span data-ttu-id="f10a6-207">Операторы bitwise JavaScript</span><span class="sxs-lookup"><span data-stu-id="f10a6-207">JavaScript bitwise operators</span></span>](https://www.w3schools.com/js/js_bitwise.asp)