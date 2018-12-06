# <a name="highresolutioniconurl-element"></a><span data-ttu-id="c6421-101">Элемент HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="c6421-101">HighResolutionIconUrl element</span></span>

<span data-ttu-id="c6421-102">Указывает URL-адрес изображения, которое используется для представления надстройки Office в пользовательском интерфейсе вставки и Магазине Office на экранах с высоким DPI.</span><span class="sxs-lookup"><span data-stu-id="c6421-102">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="c6421-103">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="c6421-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c6421-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="c6421-104">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="c6421-105">Может содержать:</span><span class="sxs-lookup"><span data-stu-id="c6421-105">Can contain</span></span>

[<span data-ttu-id="c6421-106">Переопределение</span><span class="sxs-lookup"><span data-stu-id="c6421-106">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="c6421-107">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c6421-107">Attributes</span></span>

|<span data-ttu-id="c6421-108">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="c6421-108">**Attribute**</span></span>|<span data-ttu-id="c6421-109">**Тип**</span><span class="sxs-lookup"><span data-stu-id="c6421-109">**Type**</span></span>|<span data-ttu-id="c6421-110">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="c6421-110">**Required**</span></span>|<span data-ttu-id="c6421-111">**Описание**</span><span class="sxs-lookup"><span data-stu-id="c6421-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="c6421-112">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="c6421-112">DefaultValue</span></span>|<span data-ttu-id="c6421-113">string (URL-адрес)</span><span class="sxs-lookup"><span data-stu-id="c6421-113">string (URL)</span></span>|<span data-ttu-id="c6421-114">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c6421-114">required</span></span>|<span data-ttu-id="c6421-115">Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="c6421-115">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="c6421-116">Замечания</span><span class="sxs-lookup"><span data-stu-id="c6421-116">Remarks</span></span>

<span data-ttu-id="c6421-p101">Значок почтовой надстройки отображается в разделе **Файл**  >  **Управление надстройками**. Значок надстройки области задач или контентной надстройки отображается в разделе **Вставка**  >  **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="c6421-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI . For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="c6421-119">Изображение должно быть в формате GIF, JPG, PNG, EXIF, BMP или TIFF.</span><span class="sxs-lookup"><span data-stu-id="c6421-119">The image must be in one of the following file formats at a recommended resolution of 64 x 64 pixels: GIF, JPG, PNG, EXIF, BMP or TIFF.</span></span> <span data-ttu-id="c6421-120">Для приложений области задач и приложений для работы с контентом рекомендуется размер изображения 64 х 64 пикселя.</span><span class="sxs-lookup"><span data-stu-id="c6421-120">For content and task pane apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="c6421-121">Для почтовых приложений изображение должно иметь размер 128 x 128 пикселей.</span><span class="sxs-lookup"><span data-stu-id="c6421-121">For mail apps, the image must be 128 x 128 pixels.</span></span> <span data-ttu-id="c6421-122">Дополнительные сведения см. в разделе _Создание согласованного визуального образа приложения_ статьи [Создание эффективных описаний в AppSource и в Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="c6421-122">For more information, see the section  Create a consistent visual identity for your app in Create effective Office Store apps and add-ins.</span></span>
