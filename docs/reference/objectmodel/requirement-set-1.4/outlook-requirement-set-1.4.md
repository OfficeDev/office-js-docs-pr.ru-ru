---
title: Набор обязательных элементов API для надстройки Outlook 1.4
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: 4e0d20b5449483eb3f5737fcccd0b3cd0620382a
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325201"
---
# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="44321-102">Набор обязательных элементов API для надстройки Outlook 1.4</span><span class="sxs-lookup"><span data-stu-id="44321-102">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="44321-103">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="44321-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="44321-104">В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="44321-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="44321-105">Новые возможности в версии 1.4</span><span class="sxs-lookup"><span data-stu-id="44321-105">What's new in 1.4?</span></span>

<span data-ttu-id="44321-p101">Набор обязательных элементов 1.4 включает все возможности [набора обязательных элементов версии 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). В нем добавлен доступ к пространству имен `Office.ui`.</span><span class="sxs-lookup"><span data-stu-id="44321-p101">Requirement set 1.4 includes all of the features of [Requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="44321-108">Журнал изменений</span><span class="sxs-lookup"><span data-stu-id="44321-108">Change log</span></span>

- <span data-ttu-id="44321-109">Добавлен метод [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-). Отображает диалоговое окно в ведущем приложении Office.</span><span class="sxs-lookup"><span data-stu-id="44321-109">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office host.</span></span>
- <span data-ttu-id="44321-110">Добавлен метод [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-). Доставляет сообщение из диалогового окна родительской странице.</span><span class="sxs-lookup"><span data-stu-id="44321-110">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="44321-111">Добавлен объект [Dialog](/javascript/api/office/office.dialog). Объект, возвращаемый при вызове метода [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="44321-111">Added [Dialog](/javascript/api/office/office.dialog) object: The object that is returned when the [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="44321-112">См. также</span><span class="sxs-lookup"><span data-stu-id="44321-112">See also</span></span>

- [<span data-ttu-id="44321-113">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="44321-113">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="44321-114">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="44321-114">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="44321-115">Начало работы</span><span class="sxs-lookup"><span data-stu-id="44321-115">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="44321-116">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="44321-116">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
