<span data-ttu-id="08d88-101">Outlook надстройки в основном используют API, выставленные через объект [почтовых ящиков.](/javascript/api/outlook/office.mailbox)</span><span class="sxs-lookup"><span data-stu-id="08d88-101">Outlook add-ins primarily use the APIs exposed through the [Mailbox](/javascript/api/outlook/office.mailbox) object.</span></span> <span data-ttu-id="08d88-102">Чтобы получить объекты и члены специально для использования в надстройках Outlook, такие как объект [Item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md), используйте свойство [mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) объекта **Context** для получения доступа к объекту **Mailbox**, как показано в следующей строке кода.</span><span class="sxs-lookup"><span data-stu-id="08d88-102">To access the objects and members specifically for use in Outlook add-ins, such as the [Item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) object, you use the [mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) property of the **Context** object to access the **Mailbox** object, as shown in the following line of code.</span></span>

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

<span data-ttu-id="08d88-103">Кроме того, Outlook надстройки могут использовать следующие объекты.</span><span class="sxs-lookup"><span data-stu-id="08d88-103">Additionally, Outlook add-ins can use the following objects.</span></span>

-  <span data-ttu-id="08d88-104">Объект **Office** для инициализации.</span><span class="sxs-lookup"><span data-stu-id="08d88-104">**Office** object: for initialization.</span></span>

-  <span data-ttu-id="08d88-105">Объект **Context** для получения доступа к контенту и отображения языковых свойств.</span><span class="sxs-lookup"><span data-stu-id="08d88-105">**Context** object: for access to content and display language properties.</span></span>

-  <span data-ttu-id="08d88-106">Объект **RoamingSettings** для сохранения пользовательских свойств, относящихся к надстройке Outlook, в почтовом ящике пользователя, в котором установлено приложение.</span><span class="sxs-lookup"><span data-stu-id="08d88-106">**RoamingSettings** object: for saving Outlook add-in-specific custom settings to the user's mailbox where the add-in is installed.</span></span>

<span data-ttu-id="08d88-107">Сведения об использовании API Outlook JavaScript см. в Outlook [надстройки.](../outlook/outlook-add-ins-overview.md)</span><span class="sxs-lookup"><span data-stu-id="08d88-107">For information about using the Outlook JavaScript API, see [Outlook add-ins](../outlook/outlook-add-ins-overview.md).</span></span>