# <a name="requestedheight-element"></a><span data-ttu-id="b1294-101">Элемент RequestedHeight</span><span class="sxs-lookup"><span data-stu-id="b1294-101">RequestedHeight element</span></span>

<span data-ttu-id="b1294-102">Указывает исходную высоту окна контентной или почтовой надстройки (в пикселях).</span><span class="sxs-lookup"><span data-stu-id="b1294-102">Specifies the initial height (in pixels) of a content add-in or mail add-in.</span></span> 

<span data-ttu-id="b1294-103">**Тип надстройки:** контентные и почтовые надстройки</span><span class="sxs-lookup"><span data-stu-id="b1294-103">**Add-in type:** Content, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b1294-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="b1294-104">Syntax</span></span>

```XML
<RequestedHeight>integer</RequestedHeight>
```

## <a name="contained-in"></a><span data-ttu-id="b1294-105">Родительские элементы</span><span class="sxs-lookup"><span data-stu-id="b1294-105">Contained in:</span></span>

- <span data-ttu-id="b1294-106">[DefaultSettings](defaultsettings.md) (контентные надстройки) со значением в диапазоне от 32 до 1000</span><span class="sxs-lookup"><span data-stu-id="b1294-106">[DefaultSettings](defaultsettings.md) (Content add-ins) with a value that can be between 32 and 1000</span></span>
- <span data-ttu-id="b1294-107">[DesktopSettings](desktopsettings.md) и [TabletSettings](tabletsettings.md) (почтовые надстройки) со значением в диапазоне от 32 до 450</span><span class="sxs-lookup"><span data-stu-id="b1294-107">[DesktopSettings](desktopsettings.md) and [TabletSettings](tabletsettings.md) (Mail add-ins) with a value that can be between 32 and 450</span></span>
- <span data-ttu-id="b1294-108">[ExtensionPoint](extensionpoint.md)  (контекстные почтовые надстройки) со значением в диапазоне от 140 до 450 для точки расширения **DetectedEntity** и в диапазоне от 32 до 450 для точки расширения **CustomPane**</span><span class="sxs-lookup"><span data-stu-id="b1294-108">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins) with a value that can be between 140 and 450 for the **DetectedEntity** extension point and between 32 and 450 for the **CustomPane** extension point</span></span>