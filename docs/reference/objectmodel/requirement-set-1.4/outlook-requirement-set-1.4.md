---
title: Набор обязательных элементов API для надстройки Outlook 1.4
description: Функции и API, которые были представлены для надстроек Outlook и API JavaScript для Office в составе API почтовых ящиков 1,4.
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: b4460315412e1a82473a1c33319fb960b73a5a61
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293760"
---
# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="8fd4c-103">Набор обязательных элементов API для надстройки Outlook 1.4</span><span class="sxs-lookup"><span data-stu-id="8fd4c-103">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="8fd4c-104">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="8fd4c-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="8fd4c-105">В этой документации рассматривается не последняя версия [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="8fd4c-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="8fd4c-106">Новые возможности в версии 1.4</span><span class="sxs-lookup"><span data-stu-id="8fd4c-106">What's new in 1.4?</span></span>

<span data-ttu-id="8fd4c-p101">Набор обязательных элементов 1.4 включает все возможности [набора обязательных элементов версии 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). В нем добавлен доступ к пространству имен `Office.ui`.</span><span class="sxs-lookup"><span data-stu-id="8fd4c-p101">Requirement set 1.4 includes all of the features of [Requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="8fd4c-109">Журнал изменений</span><span class="sxs-lookup"><span data-stu-id="8fd4c-109">Change log</span></span>

- <span data-ttu-id="8fd4c-110">Добавлен элемент [Office. Context. UI. displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): отображает диалоговое окно в приложении Office.</span><span class="sxs-lookup"><span data-stu-id="8fd4c-110">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office application.</span></span>
- <span data-ttu-id="8fd4c-111">Добавлен метод [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-). Доставляет сообщение из диалогового окна родительской странице.</span><span class="sxs-lookup"><span data-stu-id="8fd4c-111">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="8fd4c-112">Добавлен объект [Dialog](/javascript/api/office/office.dialog). Объект, возвращаемый при вызове метода [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="8fd4c-112">Added [Dialog](/javascript/api/office/office.dialog) object: The object that is returned when the [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="8fd4c-113">См. также</span><span class="sxs-lookup"><span data-stu-id="8fd4c-113">See also</span></span>

- [<span data-ttu-id="8fd4c-114">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="8fd4c-114">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="8fd4c-115">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="8fd4c-115">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="8fd4c-116">Начало работы</span><span class="sxs-lookup"><span data-stu-id="8fd4c-116">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="8fd4c-117">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="8fd4c-117">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
