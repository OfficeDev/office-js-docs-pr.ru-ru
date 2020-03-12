---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: ''
ms.date: 03/04/2020
localization_priority: Normal
ms.openlocfilehash: 4365dab3d8dd1ddb876536b3030926d68a89ac49
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/12/2020
ms.locfileid: "42605675"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="334b2-102">Предварительная версия набора обязательных элементов API для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="334b2-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="334b2-103">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="334b2-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="334b2-104">Эта документация относится к **предварительной версии** [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="334b2-104">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="334b2-105">Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке.</span><span class="sxs-lookup"><span data-stu-id="334b2-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="334b2-106">Не следует указывать этот набор обязательных элементов в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="334b2-106">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="334b2-107">Предварительная версия набора обязательных элементов включает все возможности [набора обязательных элементов 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="334b2-107">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="334b2-108">Возможности предварительной версии</span><span class="sxs-lookup"><span data-stu-id="334b2-108">Features in preview</span></span>

<span data-ttu-id="334b2-109">Ниже перечислены возможности предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="334b2-109">The following features are in preview.</span></span>

### <a name="append-on-send"></a><span data-ttu-id="334b2-110">Добавление при отправке</span><span class="sxs-lookup"><span data-stu-id="334b2-110">Append on send</span></span>

#### <a name="officebodyappendonsendasync"></a>[<span data-ttu-id="334b2-111">Office. Body. Аппендонсендасинк</span><span class="sxs-lookup"><span data-stu-id="334b2-111">Office.Body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="334b2-112">Добавлена новая функция для `Body` объекта, который добавляет данные в конец тела элемента в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="334b2-112">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="334b2-113">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="334b2-113">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="334b2-114">екстендедпермиссионс</span><span class="sxs-lookup"><span data-stu-id="334b2-114">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="334b2-115">Добавлен новый элемент в манифест, где `AppendOnSend` расширенное разрешение должно быть включено в коллекцию расширенных разрешений.</span><span class="sxs-lookup"><span data-stu-id="334b2-115">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="334b2-116">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="334b2-116">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="334b2-117">Взаимодействие с интерактивными сообщениями</span><span class="sxs-lookup"><span data-stu-id="334b2-117">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="334b2-118">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="334b2-118">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="334b2-119">Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="334b2-119">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="334b2-120">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="334b2-120">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="334b2-121">Тема Office</span><span class="sxs-lookup"><span data-stu-id="334b2-121">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="334b2-122">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="334b2-122">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="334b2-123">Добавлена возможность получения темы Office.</span><span class="sxs-lookup"><span data-stu-id="334b2-123">Added ability to get Office theme.</span></span>

<span data-ttu-id="334b2-124">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="334b2-124">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="334b2-125">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="334b2-125">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="334b2-126">Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="334b2-126">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="334b2-127">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="334b2-127">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="334b2-128">Единый вход</span><span class="sxs-lookup"><span data-stu-id="334b2-128">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="334b2-129">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="334b2-129">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="334b2-130">Добавлена возможность доступа к `getAccessToken`, что позволяет надстройкам [получать маркер доступа](../../../outlook/authenticate-a-user-with-an-sso-token.md) для API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="334b2-130">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="334b2-131">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="334b2-131">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="334b2-132">См. также</span><span class="sxs-lookup"><span data-stu-id="334b2-132">See also</span></span>

- [<span data-ttu-id="334b2-133">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="334b2-133">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="334b2-134">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="334b2-134">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="334b2-135">Начало работы</span><span class="sxs-lookup"><span data-stu-id="334b2-135">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="334b2-136">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="334b2-136">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
