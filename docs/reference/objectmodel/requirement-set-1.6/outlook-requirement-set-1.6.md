---
title: Набор обязательных элементов API для надстройки Outlook 1.6
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 0e1f920c259ca1ef8a137bab07132b015d9c75d2
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451733"
---
# <a name="outlook-add-in-api-requirement-set-16"></a><span data-ttu-id="08be7-102">Набор обязательных элементов API для надстройки Outlook 1.6</span><span class="sxs-lookup"><span data-stu-id="08be7-102">Outlook add-in API requirement set 1.6</span></span>

<span data-ttu-id="08be7-103">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="08be7-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="08be7-104">В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="08be7-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-16"></a><span data-ttu-id="08be7-105">Новые возможности в версии 1.6</span><span class="sxs-lookup"><span data-stu-id="08be7-105">What's new in 1.6?</span></span>

<span data-ttu-id="08be7-106">Набор обязательных элементов 1.6 включает все возможности [набора обязательных элементов версии 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span><span class="sxs-lookup"><span data-stu-id="08be7-106">Requirement set 1.6 includes all of the features of [Requirement set 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span></span> <span data-ttu-id="08be7-107">В нем добавлены перечисленные ниже возможности.</span><span class="sxs-lookup"><span data-stu-id="08be7-107">It added the following features.</span></span>

- <span data-ttu-id="08be7-108">Добавлены новые API для контекстных надстроек, которые позволяют получить соответствие объекта или RegEx, выбранного пользователем для активации надстройки.</span><span class="sxs-lookup"><span data-stu-id="08be7-108">Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.</span></span>
- <span data-ttu-id="08be7-109">Добавлен новый интерфейс API для открытия новой формы сообщения.</span><span class="sxs-lookup"><span data-stu-id="08be7-109">Added a new API to open a new message form.</span></span>
- <span data-ttu-id="08be7-110">Добавлена возможность надстройки определять тип учетной записи почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="08be7-110">Added the ability for the add-in to determine the account type of the user's mailbox.</span></span>

### <a name="change-log"></a><span data-ttu-id="08be7-111">Журнал изменений</span><span class="sxs-lookup"><span data-stu-id="08be7-111">Change log</span></span>

- <span data-ttu-id="08be7-112">Добавлен объект [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entities). Добавляет новую функцию, которая возвращает объекты, найденные в выделенном совпадении.</span><span class="sxs-lookup"><span data-stu-id="08be7-112">Added [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entities): Adds a new function that gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="08be7-113">Выделенные совпадения применяются к контекстным надстройкам.</span><span class="sxs-lookup"><span data-stu-id="08be7-113">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="08be7-114">Добавлен объект [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object). Добавляет новую функцию, которая возвращает строковые значения в выделенном совпадении, соответствующие регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="08be7-114">Added [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object): Adds a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file.</span></span> <span data-ttu-id="08be7-115">Выделенные совпадения применяются к контекстным надстройкам.</span><span class="sxs-lookup"><span data-stu-id="08be7-115">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="08be7-116">Добавлен объект [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters). Добавляет новую функцию, которая открывает новую форму сообщения.</span><span class="sxs-lookup"><span data-stu-id="08be7-116">Added [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters): Adds a new function that opens a new message form.</span></span>
- <span data-ttu-id="08be7-117">Добавлен объект [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string). Добавляет новый элемент в профиль пользователя, указывающий тип учетной записи пользователя.</span><span class="sxs-lookup"><span data-stu-id="08be7-117">Added [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string): Adds a new member to the user profile that indicates the type of the user's account.</span></span>

## <a name="see-also"></a><span data-ttu-id="08be7-118">См. также</span><span class="sxs-lookup"><span data-stu-id="08be7-118">See also</span></span>

- [<span data-ttu-id="08be7-119">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="08be7-119">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="08be7-120">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="08be7-120">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="08be7-121">Начало работы</span><span class="sxs-lookup"><span data-stu-id="08be7-121">Get started</span></span>](/outlook/add-ins/quick-start)
