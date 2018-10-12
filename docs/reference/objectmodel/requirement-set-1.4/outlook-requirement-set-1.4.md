# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="64146-101">Набор требований API для надстройки Outlook 1.4</span><span class="sxs-lookup"><span data-stu-id="64146-101">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="64146-102">Вложенный набор API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="64146-102">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="64146-103">В этой документации рассматривается не последняя версия [набора требований](/javascript/office/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="64146-103">Note: This documentation is for a [requirement set](/javascript/office/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="64146-104">Новые возможности в версии 1.4</span><span class="sxs-lookup"><span data-stu-id="64146-104">What's new in 1.4?</span></span>

<span data-ttu-id="64146-p101">Набор требований 1.4 включает все возможности [набора требований версии 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). В нем добавлен доступ к пространству имен `Office.ui`.</span><span class="sxs-lookup"><span data-stu-id="64146-p101">Requirement set 1.4 includes all of the features of [Requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="64146-107">Журнал изменений</span><span class="sxs-lookup"><span data-stu-id="64146-107">Change log</span></span>

- <span data-ttu-id="64146-108">Добавлен метод [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-). Отображает диалоговое окно в ведущем приложении Office.</span><span class="sxs-lookup"><span data-stu-id="64146-108">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office host.</span></span>
- <span data-ttu-id="64146-109">Добавлен метод [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-messageobject-). Доставляет сообщение из диалогового окна родительской странице.</span><span class="sxs-lookup"><span data-stu-id="64146-109">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-messageobject-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="64146-110">Добавлен объект[Dialog](/javascript/api/office/office.dialog): Объект, возвращаемый при вызове метода [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-).</span><span class="sxs-lookup"><span data-stu-id="64146-110">Added Dialog object: The object that is returned when the  method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="64146-111">См. также</span><span class="sxs-lookup"><span data-stu-id="64146-111">See also</span></span>

- [<span data-ttu-id="64146-112">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="64146-112">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="64146-113">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="64146-113">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="64146-114">Начало работы</span><span class="sxs-lookup"><span data-stu-id="64146-114">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)