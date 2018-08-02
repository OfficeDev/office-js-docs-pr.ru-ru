---
title: Контентные надстройки Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: f2632e94e0a797836f73caf0d53fdc0f24bd6790
ms.sourcegitcommit: bc68b4cf811b45e8b8d1cbd7c8d2867359ab671b
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2018
ms.locfileid: "21703919"
---
# <a name="content-office-add-ins"></a><span data-ttu-id="7211b-102">Контентные надстройки Office</span><span class="sxs-lookup"><span data-stu-id="7211b-102">Content Office Add-ins</span></span>

<span data-ttu-id="7211b-p101">Контентные надстройки — это рабочие области, которые можно внедрять прямо в документы Word, Excel и PowerPoint. Контентные надстройки предоставляют пользователям доступ к элементам управления интерфейсом, которые выполняют код для изменения документов или отображения данных. Используйте контентные надстройки, когда нужно внедрить функции прямо в документ.</span><span class="sxs-lookup"><span data-stu-id="7211b-p101">Content add-ins are surfaces that can be embedded directly into Word, Excel, or PowerPoint documents. Content add-ins give users access to interface controls that run code to modify documents or display data from a data source. Use content add-ins when you want to embed functionality directly into the document.</span></span>  

<span data-ttu-id="7211b-106">*Рисунок 1. Макет для контентных надстроек*</span><span class="sxs-lookup"><span data-stu-id="7211b-106">*Figure 1. Typical layout for content add-ins*</span></span>

![Изображение, на котором показан типичный макет контентной надстройки.](../images/overview-with-app-content.png)

## <a name="best-practices"></a><span data-ttu-id="7211b-108">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="7211b-108">Best practices</span></span>

- <span data-ttu-id="7211b-109">Включите элемент навигации или управления, такой как CommandBar или Pivot, в верхнюю часть надстройки.</span><span class="sxs-lookup"><span data-stu-id="7211b-109">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span>
- <span data-ttu-id="7211b-110">Включите элемент фирменной символики, такой как BrandBar, в нижнюю часть надстройки (применимо только к надстройкам Word, Excel и PowerPoint).</span><span class="sxs-lookup"><span data-stu-id="7211b-110">Include a branding element such as the BrandBar at the bottom of your add-in (applies to Word, Excel, and PowerPoint add-ins only).</span></span>

## <a name="variants"></a><span data-ttu-id="7211b-111">Варианты</span><span class="sxs-lookup"><span data-stu-id="7211b-111">Variants</span></span>

<span data-ttu-id="7211b-112">Размеры контентных надстроек для Word, Excel и PowerPoint в Office 2016 для настольных систем и Office 365 указывает пользователь.</span><span class="sxs-lookup"><span data-stu-id="7211b-112">Content add-in sizes for Word, Excel, and PowerPoint in Office 2016 desktop and Office 365 are user specified.</span></span>

## <a name="personality-menu"></a><span data-ttu-id="7211b-113">Меню личных данных</span><span class="sxs-lookup"><span data-stu-id="7211b-113">Personality menu</span></span>

<span data-ttu-id="7211b-p102">Меню личных данных могут перекрывать элементы навигации и управления, расположенные в правой верхней части надстройки. Ниже указаны текущие размеры меню личных данных в Windows и Mac.</span><span class="sxs-lookup"><span data-stu-id="7211b-p102">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="7211b-116">В Windows меню личных данных имеет размер 12 x 32 пикселей, как показано на изображении.</span><span class="sxs-lookup"><span data-stu-id="7211b-116">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="7211b-117">*Рисунок 2. Меню личных данных в Windows*</span><span class="sxs-lookup"><span data-stu-id="7211b-117">*Figure 2. Personality menu on Windows*</span></span> 

![Изображение меню личных данных на компьютере с Windows](../images/personality-menu-win.png)


<span data-ttu-id="7211b-119">В Mac меню личных данных имеет размер 26 x 26 точек, но сдвинуто на 8 пикселей влево и на 6 вниз, из-за чего оно занимает пространство размером 34 x 32 пикселей, как показано на изображении.</span><span class="sxs-lookup"><span data-stu-id="7211b-119">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the occupied space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="7211b-120">*Рисунок 3. Меню личных данных на Mac*</span><span class="sxs-lookup"><span data-stu-id="7211b-120">*Figure 3. Personality menu on Mac*</span></span>

![Изображение меню личных данных на компьютере с Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="7211b-122">Реализация</span><span class="sxs-lookup"><span data-stu-id="7211b-122">Implementation</span></span>

<span data-ttu-id="7211b-123">Пример реализации контентной надстройки для Excel: [Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) на сайте GitHub.</span><span class="sxs-lookup"><span data-stu-id="7211b-123">For a sample that implements a content add-in, see [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) in GitHub.</span></span>

## <a name="support-considerations"></a><span data-ttu-id="7211b-124">Что касается поддержки</span><span class="sxs-lookup"><span data-stu-id="7211b-124">Support considerations</span></span>
- <span data-ttu-id="7211b-125">Проверьте, будет ли ваша надстройка Office работать на [конкретной платформе Office](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability).</span><span class="sxs-lookup"><span data-stu-id="7211b-125">Check to see if your Office Add-in will work on a [specific Office host platform](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability).</span></span> 
- <span data-ttu-id="7211b-126">Для некоторого содержимого может потребоваться, чтобы пользователь добавил надстройку в список «доверенных» с тем, чтобы надстройка могла читать и записывать данные в Excel или PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="7211b-126">Some content add-ins may require the user to "trust" the add-in to read and write to Excel or PowerPoint.</span></span> <span data-ttu-id="7211b-127">Вы можете объявить нужный [уровень разрешений](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="7211b-127">You can declare what [level of permissions](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) you want your use to have in the add-in's manifest.</span></span>  
- <span data-ttu-id="7211b-128">Контентные надстройки поддерживаются в Excel и PowerPoint в Office 2013 и более поздних версий.</span><span class="sxs-lookup"><span data-stu-id="7211b-128">Content add-ins are supported in Excel and PowerPoint in Office 2013 version and later.</span></span> <span data-ttu-id="7211b-129">Если вы откроете надстройку в версии Office, которая не поддерживает веб-надстройки, вместо надстройки будет показано изображение.</span><span class="sxs-lookup"><span data-stu-id="7211b-129">If you open an add-in in a version of Office that doesn't support Office web add-ins, the add-in will be displayed as an image.</span></span>

## <a name="see-also"></a><span data-ttu-id="7211b-130">См. также</span><span class="sxs-lookup"><span data-stu-id="7211b-130">See also</span></span>
- [<span data-ttu-id="7211b-131">Сведения о доступности элементов для надстроек Office, представленные с учетом ведущих приложений и платформ</span><span class="sxs-lookup"><span data-stu-id="7211b-131">Office Add-in host and platform availability</span></span>](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-in-availability)
- [<span data-ttu-id="7211b-132">Office UI Fabric в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="7211b-132">Office UI Fabric in Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/design/office-ui-fabric) 
- [<span data-ttu-id="7211b-133">Шаблоны проектирования взаимодействия для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="7211b-133">UX design patterns for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/design/ux-design-pattern-templates)
- [<span data-ttu-id="7211b-134">Запрашивание разрешений на использование API в контентных надстройках и надстройках области задач</span><span class="sxs-lookup"><span data-stu-id="7211b-134">Requesting permissions for API use in content and task pane add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)
