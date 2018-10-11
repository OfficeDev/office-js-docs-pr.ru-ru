# <a name="metadata-element"></a><span data-ttu-id="bdc7c-101">Элемент Metadata</span><span class="sxs-lookup"><span data-stu-id="bdc7c-101">MetaData element</span></span>

<span data-ttu-id="bdc7c-102">Задает параметры метаданных, используемых настраиваемыми функциями в  Excel.</span><span class="sxs-lookup"><span data-stu-id="bdc7c-102">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="bdc7c-103">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="bdc7c-103">Attributes</span></span>

<span data-ttu-id="bdc7c-104">Нет</span><span class="sxs-lookup"><span data-stu-id="bdc7c-104">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="bdc7c-105">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="bdc7c-105">Child elements</span></span>

|  <span data-ttu-id="bdc7c-106">Элемент</span><span class="sxs-lookup"><span data-stu-id="bdc7c-106">Element</span></span>  |  <span data-ttu-id="bdc7c-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="bdc7c-107">Required</span></span>  |  <span data-ttu-id="bdc7c-108">Описание</span><span class="sxs-lookup"><span data-stu-id="bdc7c-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="bdc7c-109">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="bdc7c-109">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="bdc7c-110">Да</span><span class="sxs-lookup"><span data-stu-id="bdc7c-110">Yes</span></span>  | <span data-ttu-id="bdc7c-111">Строка с идентификатором ресурса файла HTML, используемого настраиваемыми функциями.</span><span class="sxs-lookup"><span data-stu-id="bdc7c-111">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="bdc7c-112">Пример</span><span class="sxs-lookup"><span data-stu-id="bdc7c-112">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
