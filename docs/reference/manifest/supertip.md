# <a name="supertip"></a><span data-ttu-id="8c95f-101">Supertip</span><span class="sxs-lookup"><span data-stu-id="8c95f-101">Supertip</span></span>

<span data-ttu-id="8c95f-p101">Определяет форматированную подсказку (элементы Title и Description). Используется элементами управления [Кнопка](control.md#button-control) или [Меню](control.md#menu-dropdown-button-controls).</span><span class="sxs-lookup"><span data-stu-id="8c95f-p101">Defines a rich tooltip (both Title and Description). It is used by both [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls)  controls.</span></span>

## <a name="child-elements"></a><span data-ttu-id="8c95f-104">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="8c95f-104">Child elements</span></span>

|  <span data-ttu-id="8c95f-105">Элемент</span><span class="sxs-lookup"><span data-stu-id="8c95f-105">Element</span></span> |  <span data-ttu-id="8c95f-106">Обязательный</span><span class="sxs-lookup"><span data-stu-id="8c95f-106">Required</span></span>  |  <span data-ttu-id="8c95f-107">Description</span><span class="sxs-lookup"><span data-stu-id="8c95f-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="8c95f-108">Title</span><span class="sxs-lookup"><span data-stu-id="8c95f-108">Title</span></span>](#title)        | <span data-ttu-id="8c95f-109">Да</span><span class="sxs-lookup"><span data-stu-id="8c95f-109">Yes</span></span> |   <span data-ttu-id="8c95f-110">Текст суперподсказки.</span><span class="sxs-lookup"><span data-stu-id="8c95f-110">The text for the supertip.</span></span>         |
|  [<span data-ttu-id="8c95f-111">Description</span><span class="sxs-lookup"><span data-stu-id="8c95f-111">Description</span></span>](#description)  | <span data-ttu-id="8c95f-112">Да</span><span class="sxs-lookup"><span data-stu-id="8c95f-112">Yes</span></span> |  <span data-ttu-id="8c95f-113">Описание суперподсказки.</span><span class="sxs-lookup"><span data-stu-id="8c95f-113">The description for the supertip.</span></span>    |

### <a name="title"></a><span data-ttu-id="8c95f-114">Название</span><span class="sxs-lookup"><span data-stu-id="8c95f-114">Title</span></span>

<span data-ttu-id="8c95f-p102">Обязательный элемент. Текст суперподсказки. Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **String** в элементе **ShortStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="8c95f-p102">Required. The text for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **ShortStrings** element in the [Resources](resources.md) element.</span></span>

### <a name="description"></a><span data-ttu-id="8c95f-118">Description</span><span class="sxs-lookup"><span data-stu-id="8c95f-118">Description</span></span>

<span data-ttu-id="8c95f-p103">Обязательный элемент. Описание суперподсказки. Атрибуту **resid** должно быть присвоено значение атрибута **id** элемента **String** в элементе **LongStrings**, вложенном в элемент [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="8c95f-p103">Required. The description for the supertip. The  **resid** attribute must be set to the value of the **id** attribute of a **String** element in the **LongStrings** element in the [Resources](resources.md) element.</span></span>

## <a name="example"></a><span data-ttu-id="8c95f-122">Пример</span><span class="sxs-lookup"><span data-stu-id="8c95f-122">Example</span></span>

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```
