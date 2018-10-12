# <a name="sourcelocation-element"></a><span data-ttu-id="af0dc-101">Элемент SourceLocation</span><span class="sxs-lookup"><span data-stu-id="af0dc-101">SourceLocation element</span></span>

<span data-ttu-id="af0dc-p101">Указывает расположения исходного файла для надстройки Office как URL-адреса длиной от 1 до 2018 символов. В качестве источника необходимо указать адрес HTTPS, а не путь к файлу.</span><span class="sxs-lookup"><span data-stu-id="af0dc-p101">Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="af0dc-104">**Тип надстройки:** содержимое, область задач, почта</span><span class="sxs-lookup"><span data-stu-id="af0dc-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="af0dc-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="af0dc-105">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="af0dc-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="af0dc-106">Contained in:</span></span>

- <span data-ttu-id="af0dc-107">[DefaultSettings](defaultsettings.md) (надстройки области задач и контентные надстройки)</span><span class="sxs-lookup"><span data-stu-id="af0dc-107">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="af0dc-108">[FormSettings](formsettings.md) (почтовые надстройки)</span><span class="sxs-lookup"><span data-stu-id="af0dc-108">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="af0dc-109">[ExtensionPoint](extensionpoint.md) (контекстные почтовые надстройки)</span><span class="sxs-lookup"><span data-stu-id="af0dc-109">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="af0dc-110">Может содержать</span><span class="sxs-lookup"><span data-stu-id="af0dc-110">Can contain:</span></span>

[<span data-ttu-id="af0dc-111">Переопределение</span><span class="sxs-lookup"><span data-stu-id="af0dc-111">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="af0dc-112">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="af0dc-112">Attributes</span></span>

|<span data-ttu-id="af0dc-113">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="af0dc-113">**Attribute**</span></span>|<span data-ttu-id="af0dc-114">**Тип**</span><span class="sxs-lookup"><span data-stu-id="af0dc-114">**Type**</span></span>|<span data-ttu-id="af0dc-115">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="af0dc-115">**Required**</span></span>|<span data-ttu-id="af0dc-116">**Описание**</span><span class="sxs-lookup"><span data-stu-id="af0dc-116">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="af0dc-117">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="af0dc-117">DefaultValue</span></span>|<span data-ttu-id="af0dc-118">URL</span><span class="sxs-lookup"><span data-stu-id="af0dc-118">URL</span></span>|<span data-ttu-id="af0dc-119">обязательный</span><span class="sxs-lookup"><span data-stu-id="af0dc-119">required</span></span>|<span data-ttu-id="af0dc-120">Задает значение этого параметра по умолчанию для языкового стандарта, указанного в элементе [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="af0dc-120">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
