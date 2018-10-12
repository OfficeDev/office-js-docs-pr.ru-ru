# <a name="highresolutioniconurl-element"></a><span data-ttu-id="7a656-101">Элемент HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="7a656-101">HighResolutionIconUrl element</span></span>

<span data-ttu-id="7a656-102">Указывает URL-адрес изображения, которое используется для представления надстройки Office в пользовательском интерфейсе вставки и Магазине Office на экранах с высоким DPI.</span><span class="sxs-lookup"><span data-stu-id="7a656-102">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="7a656-103">**Тип надстройки:** контентные и почтовые надстройки, а также надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="7a656-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="7a656-104">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="7a656-104">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="7a656-105">Может содержать</span><span class="sxs-lookup"><span data-stu-id="7a656-105">Can contain:</span></span>

[<span data-ttu-id="7a656-106">Переопределение</span><span class="sxs-lookup"><span data-stu-id="7a656-106">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="7a656-107">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="7a656-107">Attributes</span></span>

|<span data-ttu-id="7a656-108">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="7a656-108">**Attribute**</span></span>|<span data-ttu-id="7a656-109">**Тип**</span><span class="sxs-lookup"><span data-stu-id="7a656-109">**Type**</span></span>|<span data-ttu-id="7a656-110">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="7a656-110">**Required**</span></span>|<span data-ttu-id="7a656-111">**Описание**</span><span class="sxs-lookup"><span data-stu-id="7a656-111">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="7a656-112">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="7a656-112">DefaultValue</span></span>|<span data-ttu-id="7a656-113">string (URL)</span><span class="sxs-lookup"><span data-stu-id="7a656-113">string (URL)</span></span>|<span data-ttu-id="7a656-114">обязательный</span><span class="sxs-lookup"><span data-stu-id="7a656-114">required</span></span>|<span data-ttu-id="7a656-115">Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="7a656-115">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="7a656-116">Замечания</span><span class="sxs-lookup"><span data-stu-id="7a656-116">Remarks</span></span>

<span data-ttu-id="7a656-p101">Значок почтовой надстройки отображается в разделе **Файл**  >  **Управление надстройками**. Значок надстройки области задач или контентной надстройки отображается в разделе **Вставка**  >  **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="7a656-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI . For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="7a656-119">Рекомендуемое разрешение изображения — 64 x 64 пикселя. Поддерживаемые форматы: GIF, JPG, PNG, EXIF, BMP и TIFF.</span><span class="sxs-lookup"><span data-stu-id="7a656-119">The image must be in one of the following file formats at a recommended resolution of 64 x 64 pixels: GIF, JPG, PNG, EXIF, BMP or TIFF.</span></span> <span data-ttu-id="7a656-120">Подробнее см. раздел _Создание согласованного визуального образа приложения_ в статье [Создание эффективных профилей в AppSource и Office](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings).</span><span class="sxs-lookup"><span data-stu-id="7a656-120">For more information, see the section  Create a consistent visual identity for your app in Create effective Office Store apps and add-ins.</span></span>
