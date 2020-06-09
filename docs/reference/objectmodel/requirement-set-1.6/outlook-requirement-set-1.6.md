---
title: Набор обязательных элементов API для надстройки Outlook 1.6
description: Функции и API, которые были представлены для надстроек Outlook и API JavaScript для Office в составе API почтовых ящиков 1,6.
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: 024b5ab992b146a1958653c38941434da00e1a03
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611451"
---
# <a name="outlook-add-in-api-requirement-set-16"></a><span data-ttu-id="e0aad-103">Набор обязательных элементов API для надстройки Outlook 1.6</span><span class="sxs-lookup"><span data-stu-id="e0aad-103">Outlook add-in API requirement set 1.6</span></span>

<span data-ttu-id="e0aad-104">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="e0aad-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="e0aad-105">В этой документации рассматривается не последняя версия [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="e0aad-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-16"></a><span data-ttu-id="e0aad-106">Новые возможности в версии 1.6</span><span class="sxs-lookup"><span data-stu-id="e0aad-106">What's new in 1.6?</span></span>

<span data-ttu-id="e0aad-107">Набор обязательных элементов 1.6 включает все возможности [набора обязательных элементов версии 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span><span class="sxs-lookup"><span data-stu-id="e0aad-107">Requirement set 1.6 includes all of the features of [Requirement set 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span></span> <span data-ttu-id="e0aad-108">В нем добавлены перечисленные ниже возможности.</span><span class="sxs-lookup"><span data-stu-id="e0aad-108">It added the following features.</span></span>

- <span data-ttu-id="e0aad-109">Добавлены новые API для контекстных надстроек, которые позволяют получить соответствие объекта или RegEx, выбранного пользователем для активации надстройки.</span><span class="sxs-lookup"><span data-stu-id="e0aad-109">Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.</span></span>
- <span data-ttu-id="e0aad-110">Добавлен новый интерфейс API для открытия новой формы сообщения.</span><span class="sxs-lookup"><span data-stu-id="e0aad-110">Added a new API to open a new message form.</span></span>
- <span data-ttu-id="e0aad-111">Добавлена возможность надстройки определять тип учетной записи почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="e0aad-111">Added the ability for the add-in to determine the account type of the user's mailbox.</span></span>

### <a name="change-log"></a><span data-ttu-id="e0aad-112">Журнал изменений</span><span class="sxs-lookup"><span data-stu-id="e0aad-112">Change log</span></span>

- <span data-ttu-id="e0aad-113">Добавлен объект [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods). Добавляет новую функцию, которая возвращает объекты, найденные в выделенном совпадении.</span><span class="sxs-lookup"><span data-stu-id="e0aad-113">Added [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods): Adds a new function that gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="e0aad-114">Выделенные совпадения применяются к контекстным надстройкам.</span><span class="sxs-lookup"><span data-stu-id="e0aad-114">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="e0aad-115">Добавлен объект [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods). Добавляет новую функцию, которая возвращает строковые значения в выделенном совпадении, соответствующие регулярным выражениям, определенным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="e0aad-115">Added [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods): Adds a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file.</span></span> <span data-ttu-id="e0aad-116">Выделенные совпадения применяются к контекстным надстройкам.</span><span class="sxs-lookup"><span data-stu-id="e0aad-116">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="e0aad-117">Добавлен объект [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods). Добавляет новую функцию, которая открывает новую форму сообщения.</span><span class="sxs-lookup"><span data-stu-id="e0aad-117">Added [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods): Adds a new function that opens a new message form.</span></span>
- <span data-ttu-id="e0aad-118">Добавлен объект [Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6#accounttype). Добавляет новый элемент в профиль пользователя, указывающий тип учетной записи пользователя.</span><span class="sxs-lookup"><span data-stu-id="e0aad-118">Added [Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6#accounttype): Adds a new member to the user profile that indicates the type of the user's account.</span></span>

## <a name="see-also"></a><span data-ttu-id="e0aad-119">См. также</span><span class="sxs-lookup"><span data-stu-id="e0aad-119">See also</span></span>

- [<span data-ttu-id="e0aad-120">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="e0aad-120">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="e0aad-121">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="e0aad-121">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="e0aad-122">Начало работы</span><span class="sxs-lookup"><span data-stu-id="e0aad-122">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="e0aad-123">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="e0aad-123">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
