---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: ''
ms.date: 10/30/2019
localization_priority: Priority
ms.openlocfilehash: bf8f140e893a19a4dec717b985f3bbf4226db9d5
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902118"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="d52bd-102">Предварительная версия набора обязательных элементов API для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="d52bd-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="d52bd-103">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="d52bd-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d52bd-104">Эта документация относится к **предварительной версии** [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="d52bd-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="d52bd-105">Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке.</span><span class="sxs-lookup"><span data-stu-id="d52bd-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="d52bd-106">Не следует указывать этот набор обязательных элементов в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="d52bd-106">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="d52bd-107">Предварительная версия набора обязательных элементов включает все возможности [набора обязательных элементов 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="d52bd-107">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="d52bd-108">Возможности предварительной версии</span><span class="sxs-lookup"><span data-stu-id="d52bd-108">Features in preview</span></span>

<span data-ttu-id="d52bd-109">Ниже перечислены возможности предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="d52bd-109">The following features are in preview.</span></span>

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="d52bd-110">Взаимодействие с интерактивными сообщениями</span><span class="sxs-lookup"><span data-stu-id="d52bd-110">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="d52bd-111">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="d52bd-111">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="d52bd-112">Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="d52bd-112">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="d52bd-113">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="d52bd-113">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="d52bd-114">Тема Office</span><span class="sxs-lookup"><span data-stu-id="d52bd-114">Office theme</span></span>

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[<span data-ttu-id="d52bd-115">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="d52bd-115">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="d52bd-116">Добавлена возможность получения темы Office.</span><span class="sxs-lookup"><span data-stu-id="d52bd-116">Added ability to get Office theme.</span></span>

<span data-ttu-id="d52bd-117">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d52bd-117">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="d52bd-118">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="d52bd-118">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="d52bd-119">Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="d52bd-119">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="d52bd-120">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d52bd-120">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="d52bd-121">Единый вход</span><span class="sxs-lookup"><span data-stu-id="d52bd-121">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="d52bd-122">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d52bd-122">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="d52bd-123">Добавлена возможность доступа к `getAccessTokenAsync`, что позволяет надстройкам [получать маркер доступа](/outlook/add-ins/authenticate-a-user-with-an-sso-token) для API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="d52bd-123">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="d52bd-124">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="d52bd-124">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="d52bd-125">См. также</span><span class="sxs-lookup"><span data-stu-id="d52bd-125">See also</span></span>

- [<span data-ttu-id="d52bd-126">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="d52bd-126">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="d52bd-127">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="d52bd-127">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="d52bd-128">Начало работы</span><span class="sxs-lookup"><span data-stu-id="d52bd-128">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="d52bd-129">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="d52bd-129">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
