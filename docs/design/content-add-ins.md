---
title: Контентные надстройки Office
description: Контентные надстройки — это рабочие области, которые можно внедрять прямо в документы Excel или PowerPoint, что предоставляет пользователям доступ к элементам управления интерфейсом, которые выполняют код для изменения документов или отображения данных.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: f228ae8e7cca0426b0b43e31e38454029e4c7614
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093849"
---
# <a name="content-office-add-ins"></a><span data-ttu-id="aa725-103">Контентные надстройки Office</span><span class="sxs-lookup"><span data-stu-id="aa725-103">Content Office Add-ins</span></span>

<span data-ttu-id="aa725-104">Контентные надстройки — это рабочие области, которые можно внедрять прямо в документы Excel или PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="aa725-104">Content add-ins are surfaces that can be embedded directly into Excel or PowerPoint documents.</span></span> <span data-ttu-id="aa725-105">Контентные надстройки предоставляют пользователям доступ к элементам управления интерфейсом, которые выполняют код для изменения документов или отображения данных.</span><span class="sxs-lookup"><span data-stu-id="aa725-105">Content add-ins give users access to interface controls that run code to modify documents or display data from a data source.</span></span> <span data-ttu-id="aa725-106">Используйте контентные надстройки, когда требуется внедрить функции непосредственно в документ.</span><span class="sxs-lookup"><span data-stu-id="aa725-106">Use content add-ins when you want to embed functionality directly into the document.</span></span>  

<span data-ttu-id="aa725-107">*Рисунок 1. Макет для контентных надстроек*</span><span class="sxs-lookup"><span data-stu-id="aa725-107">*Figure 1. Typical layout for content add-ins*</span></span>

![Изображение, на котором показан типичный макет контентной надстройки.](../images/overview-with-app-content.png)

## <a name="best-practices"></a><span data-ttu-id="aa725-109">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="aa725-109">Best practices</span></span>

- <span data-ttu-id="aa725-110">Добавьте элемент навигации или управления, такой как CommandBar или Pivot, в верхнюю часть надстройки.</span><span class="sxs-lookup"><span data-stu-id="aa725-110">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span>
- <span data-ttu-id="aa725-111">Добавьте элемент фирменной символики, такой как BrandBar, в нижнюю часть надстройки (применимо только к надстройкам Excel и PowerPoint).</span><span class="sxs-lookup"><span data-stu-id="aa725-111">Include a branding element such as the BrandBar at the bottom of your add-in (applies to Excel and PowerPoint add-ins only).</span></span>

## <a name="variants"></a><span data-ttu-id="aa725-112">Варианты</span><span class="sxs-lookup"><span data-stu-id="aa725-112">Variants</span></span>

<span data-ttu-id="aa725-113">Размеры контентных надстроек для Excel и PowerPoint в Office для настольных ПК и Microsoft 365 указаны пользователем.</span><span class="sxs-lookup"><span data-stu-id="aa725-113">Content add-in sizes for Excel and PowerPoint in Office desktop and Microsoft 365 are user specified.</span></span>

## <a name="personality-menu"></a><span data-ttu-id="aa725-114">Меню личных данных</span><span class="sxs-lookup"><span data-stu-id="aa725-114">Personality menu</span></span>

<span data-ttu-id="aa725-115">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in.</span><span class="sxs-lookup"><span data-stu-id="aa725-115">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in.</span></span> <span data-ttu-id="aa725-116">The following are the current dimensions of the personality menu on Windows and Mac.</span><span class="sxs-lookup"><span data-stu-id="aa725-116">The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="aa725-117">В Windows меню личных данных имеет размер 12 x 32 пикселей, как показано на изображении.</span><span class="sxs-lookup"><span data-stu-id="aa725-117">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="aa725-118">*Рисунок 2. Меню личных данных в Windows*</span><span class="sxs-lookup"><span data-stu-id="aa725-118">*Figure 2. Personality menu on Windows*</span></span> 

![Изображение меню личных данных на компьютере с Windows](../images/personality-menu-win.png)


<span data-ttu-id="aa725-120">В Mac меню личных данных имеет размер 26 x 26 точек, но сдвинуто на 8 пикселей влево и на 6 вниз, из-за чего оно занимает пространство размером 34 x 32 пикселей, как показано на изображении.</span><span class="sxs-lookup"><span data-stu-id="aa725-120">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the occupied space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="aa725-121">*Рисунок 3. Меню личных данных на Mac*</span><span class="sxs-lookup"><span data-stu-id="aa725-121">*Figure 3. Personality menu on Mac*</span></span>

![Изображение меню личных данных на компьютере с Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="aa725-123">Реализация</span><span class="sxs-lookup"><span data-stu-id="aa725-123">Implementation</span></span>

<span data-ttu-id="aa725-124">Пример реализации контентной надстройки для Excel: [Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) на сайте GitHub.</span><span class="sxs-lookup"><span data-stu-id="aa725-124">For a sample that implements a content add-in, see [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) in GitHub.</span></span>

## <a name="support-considerations"></a><span data-ttu-id="aa725-125">Что касается поддержки</span><span class="sxs-lookup"><span data-stu-id="aa725-125">Support considerations</span></span>

- <span data-ttu-id="aa725-126">Проверьте, будет ли ваша надстройка Office работать на [конкретной платформе Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="aa725-126">Check to see if your Office Add-in will work on a [specific Office host platform](../overview/office-add-in-availability.md).</span></span>
- <span data-ttu-id="aa725-127">Чтобы надстройка могла читать и записывать данные в Excel или PowerPoint, может потребоваться добавление в список доверенных.</span><span class="sxs-lookup"><span data-stu-id="aa725-127">Some content add-ins may require the user to "trust" the add-in to read and write to Excel or PowerPoint.</span></span> <span data-ttu-id="aa725-128">Вы можете объявить нужный [уровень разрешений](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) для пользователя в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="aa725-128">You can declare what [level of permissions](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) you want your user to have in the add-in's manifest.</span></span>  
- <span data-ttu-id="aa725-129">Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later.</span><span class="sxs-lookup"><span data-stu-id="aa725-129">Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later.</span></span> <span data-ttu-id="aa725-130">If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.</span><span class="sxs-lookup"><span data-stu-id="aa725-130">If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.</span></span>

## <a name="see-also"></a><span data-ttu-id="aa725-131">См. также</span><span class="sxs-lookup"><span data-stu-id="aa725-131">See also</span></span>

- [<span data-ttu-id="aa725-132">Сведения о доступности элементов для надстроек Office, представленные с учетом ведущих приложений и платформ</span><span class="sxs-lookup"><span data-stu-id="aa725-132">Office Add-in host and platform availability</span></span>](../overview/office-add-in-availability.md)
- [<span data-ttu-id="aa725-133">Office UI Fabric в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="aa725-133">Office UI Fabric in Office Add-ins</span></span>](../design/office-ui-fabric.md)
- [<span data-ttu-id="aa725-134">Конструктивные шаблоны для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="aa725-134">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
- [<span data-ttu-id="aa725-135">Запрос разрешений на использование API в надстройках</span><span class="sxs-lookup"><span data-stu-id="aa725-135">Requesting permissions for API use in add-ins</span></span>](../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
