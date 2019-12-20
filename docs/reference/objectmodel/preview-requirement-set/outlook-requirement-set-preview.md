---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: ''
ms.date: 12/17/2019
localization_priority: Priority
ms.openlocfilehash: a3cc49562add2f6fe54cf83d2f2ed64ebb61d8c7
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/19/2019
ms.locfileid: "40815048"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="3bbd8-102">Предварительная версия набора обязательных элементов API для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="3bbd8-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="3bbd8-103">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="3bbd8-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3bbd8-104">Эта документация относится к **предварительной версии** [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="3bbd8-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="3bbd8-105">Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке.</span><span class="sxs-lookup"><span data-stu-id="3bbd8-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="3bbd8-106">Не следует указывать этот набор обязательных элементов в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="3bbd8-106">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="3bbd8-107">Предварительная версия набора обязательных элементов включает все возможности [набора обязательных элементов 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="3bbd8-107">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="3bbd8-108">Возможности предварительной версии</span><span class="sxs-lookup"><span data-stu-id="3bbd8-108">Features in preview</span></span>

<span data-ttu-id="3bbd8-109">Ниже перечислены возможности предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="3bbd8-109">The following features are in preview.</span></span>

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="3bbd8-110">Взаимодействие с интерактивными сообщениями</span><span class="sxs-lookup"><span data-stu-id="3bbd8-110">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdmethods"></a>[<span data-ttu-id="3bbd8-111">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="3bbd8-111">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="3bbd8-112">Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="3bbd8-112">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="3bbd8-113">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="3bbd8-113">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="3bbd8-114">Тема Office</span><span class="sxs-lookup"><span data-stu-id="3bbd8-114">Office theme</span></span>

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[<span data-ttu-id="3bbd8-115">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="3bbd8-115">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="3bbd8-116">Добавлена возможность получения темы Office.</span><span class="sxs-lookup"><span data-stu-id="3bbd8-116">Added ability to get Office theme.</span></span>

<span data-ttu-id="3bbd8-117">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3bbd8-117">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="3bbd8-118">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="3bbd8-118">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="3bbd8-119">Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="3bbd8-119">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="3bbd8-120">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3bbd8-120">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="3bbd8-121">Единый вход</span><span class="sxs-lookup"><span data-stu-id="3bbd8-121">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstokenofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="3bbd8-122">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="3bbd8-122">OfficeRuntime.auth.getAccessToken</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="3bbd8-123">Добавлена возможность доступа к `getAccessToken`, что позволяет надстройкам [получать маркер доступа](/outlook/add-ins/authenticate-a-user-with-an-sso-token) для API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="3bbd8-123">Added access to `getAccessToken`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="3bbd8-124">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="3bbd8-124">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="3bbd8-125">См. также</span><span class="sxs-lookup"><span data-stu-id="3bbd8-125">See also</span></span>

- [<span data-ttu-id="3bbd8-126">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="3bbd8-126">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="3bbd8-127">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="3bbd8-127">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="3bbd8-128">Начало работы</span><span class="sxs-lookup"><span data-stu-id="3bbd8-128">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="3bbd8-129">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="3bbd8-129">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
