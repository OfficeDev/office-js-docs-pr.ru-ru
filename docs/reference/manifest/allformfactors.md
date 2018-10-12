# <a name="allformfactors-element"></a><span data-ttu-id="6cf9d-101">Элемент AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="6cf9d-101">AllFormFactors element</span></span>

<span data-ttu-id="6cf9d-102">Указывает параметры всех форм-факторов для надстройки.</span><span class="sxs-lookup"><span data-stu-id="6cf9d-102">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="6cf9d-103">В настоящее время только в настраиваемых функциях применяется **AllFormFactors**.</span><span class="sxs-lookup"><span data-stu-id="6cf9d-103">Currently, the only feature using AllFormFactors is custom functions.</span></span> <span data-ttu-id="6cf9d-104">Элемент\*\* AllFormFactors\*\* является обязательным при использовании настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="6cf9d-104">AllFormFactors is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="6cf9d-105">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="6cf9d-105">Child elements</span></span>

|  <span data-ttu-id="6cf9d-106">Элемент</span><span class="sxs-lookup"><span data-stu-id="6cf9d-106">Element</span></span> |  <span data-ttu-id="6cf9d-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="6cf9d-107">Required</span></span>  |  <span data-ttu-id="6cf9d-108">Описание</span><span class="sxs-lookup"><span data-stu-id="6cf9d-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="6cf9d-109">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="6cf9d-109">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="6cf9d-110">Да</span><span class="sxs-lookup"><span data-stu-id="6cf9d-110">Yes</span></span> |  <span data-ttu-id="6cf9d-111">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="6cf9d-111">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="6cf9d-112">Пример использования AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="6cf9d-112">AllFormFactors example</span></span>

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
