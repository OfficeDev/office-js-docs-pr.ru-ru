---
title: Набор обязательных элементов API для надстройки Outlook 1.6
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 22702448b82a108c401f9f81d3b8a321e14ead63
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814663"
---
# <a name="outlook-add-in-api-requirement-set-16"></a><span data-ttu-id="422ae-102">Набор обязательных элементов API для надстройки Outlook 1.6</span><span class="sxs-lookup"><span data-stu-id="422ae-102">Outlook add-in API requirement set 1.6</span></span>

<span data-ttu-id="422ae-103">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="422ae-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="422ae-104">В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="422ae-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-16"></a><span data-ttu-id="422ae-105">Новые возможности в версии 1.6</span><span class="sxs-lookup"><span data-stu-id="422ae-105">What's new in 1.6?</span></span>

<span data-ttu-id="422ae-106">Набор обязательных элементов 1.6 включает все возможности [набора обязательных элементов версии 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span><span class="sxs-lookup"><span data-stu-id="422ae-106">Requirement set 1.6 includes all of the features of [Requirement set 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span></span> <span data-ttu-id="422ae-107">В нем добавлены перечисленные ниже возможности.</span><span class="sxs-lookup"><span data-stu-id="422ae-107">It added the following features.</span></span>

- <span data-ttu-id="422ae-108">Добавлены новые API для контекстных надстроек, которые позволяют получить соответствие объекта или RegEx, выбранного пользователем для активации надстройки.</span><span class="sxs-lookup"><span data-stu-id="422ae-108">Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.</span></span>
- <span data-ttu-id="422ae-109">Добавлен новый интерфейс API для открытия новой формы сообщения.</span><span class="sxs-lookup"><span data-stu-id="422ae-109">Added a new API to open a new message form.</span></span>
- <span data-ttu-id="422ae-110">Добавлена возможность надстройки определять тип учетной записи почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="422ae-110">Added the ability for the add-in to determine the account type of the user's mailbox.</span></span>

### <a name="change-log"></a><span data-ttu-id="422ae-111">Журнал изменений</span><span class="sxs-lookup"><span data-stu-id="422ae-111">Change log</span></span>

- <span data-ttu-id="422ae-112">Добавлен объект [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods). Добавляет новую функцию, которая возвращает объекты, найденные в выделенном совпадении.</span><span class="sxs-lookup"><span data-stu-id="422ae-112">Added [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods): Adds a new function that gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="422ae-113">Выделенные совпадения применяются к контекстным надстройкам.</span><span class="sxs-lookup"><span data-stu-id="422ae-113">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="422ae-114">Добавлен объект [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods). Добавляет новую функцию, которая возвращает строковые значения в выделенном совпадении, соответствующие регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="422ae-114">Added [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods): Adds a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file.</span></span> <span data-ttu-id="422ae-115">Выделенные совпадения применяются к контекстным надстройкам.</span><span class="sxs-lookup"><span data-stu-id="422ae-115">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="422ae-116">Добавлен объект [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods). Добавляет новую функцию, которая открывает новую форму сообщения.</span><span class="sxs-lookup"><span data-stu-id="422ae-116">Added [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods): Adds a new function that opens a new message form.</span></span>
- <span data-ttu-id="422ae-117">Добавлен объект [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#properties). Добавляет новый элемент в профиль пользователя, указывающий тип учетной записи пользователя.</span><span class="sxs-lookup"><span data-stu-id="422ae-117">Added [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#properties): Adds a new member to the user profile that indicates the type of the user's account.</span></span>

## <a name="see-also"></a><span data-ttu-id="422ae-118">См. также</span><span class="sxs-lookup"><span data-stu-id="422ae-118">See also</span></span>

- [<span data-ttu-id="422ae-119">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="422ae-119">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="422ae-120">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="422ae-120">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="422ae-121">Начало работы</span><span class="sxs-lookup"><span data-stu-id="422ae-121">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="422ae-122">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="422ae-122">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
