---
title: Инструкции. Использование тем документов в надстройках для PowerPoint
description: ''
ms.date: 10/14/2019
localization_priority: Normal
ms.openlocfilehash: 83b4c2192ba3c01deedfe69a8338265fbf7eaf53
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324921"
---
# <a name="use-document-themes-in-your-powerpoint-add-ins"></a><span data-ttu-id="935ef-102">Использование тем документов в надстройках PowerPoint</span><span class="sxs-lookup"><span data-stu-id="935ef-102">Use document themes in your PowerPoint add-ins</span></span>

<span data-ttu-id="935ef-p101">[Тема Office](https://support.office.com/Article/What-is-a-theme--7528ccc2-4327-4692-8bf5-9b5a3f2a5ef5) состоит из визуально согласованного набора шрифтов и цветов, которые можно применять к презентациям, документам, электронным таблицам и письмам. Чтобы применить или настроить тему презентации в PowerPoint, используйте группы **Темы** и **Варианты** на вкладке **Дизайн**. По умолчанию PowerPoint присваивает новой пустой презентации **тему Office**, но вы можете выбрать другие темы, доступные на вкладке **Дизайн**, скачать дополнительные темы с веб-сайта Office.com или создать и настроить собственную тему.</span><span class="sxs-lookup"><span data-stu-id="935ef-p101">An [Office theme](https://support.office.com/Article/What-is-a-theme--7528ccc2-4327-4692-8bf5-9b5a3f2a5ef5) consists, in part, of a visually coordinated set of fonts and colors that you can apply to presentations, documents, worksheets, and emails. To apply or customize the theme of a presentation in PowerPoint, you use the **Themes** and **Variants** groups on **Design** tab of the ribbon. PowerPoint assigns a new blank presentation with the default **Office Theme**, but you can choose other themes available on the **Design** tab, download additional themes from Office.com, or create and customize your own theme.</span></span>

<span data-ttu-id="935ef-106">Используя файл OfficeThemes.css, можно создавать надстройки, согласованные с PowerPoint, двумя способами:</span><span class="sxs-lookup"><span data-stu-id="935ef-106">Using OfficeThemes.css, helps you design add-ins that are coordinated with PowerPoint in two ways:</span></span>

- <span data-ttu-id="935ef-p102">**Контентные надстройки для PowerPoint**. Укажите шрифты и цвета, соответствующие теме презентации контентной надстройки, используя классы OfficeThemes.css для темы документа. Эти шрифты и цвета будут динамически обновляться при изменении или настройке темы презентации.</span><span class="sxs-lookup"><span data-stu-id="935ef-p102">**In content add-ins for PowerPoint**. Use the document theme classes of OfficeThemes.css to specify fonts and colors that match the theme of the presentation your content add-in is inserted into - and those fonts and colors will dynamically update if a user changes or customizes the presentation's theme.</span></span>
    
- <span data-ttu-id="935ef-p103">**Надстройки области задач для PowerPoint**. Укажите шрифты и фоновые цвета, используемые в пользовательском интерфейсе, используя классы OfficeThemes.css для темы пользовательского интерфейса Office, чтобы цвета ваших надстроек области задач соответствовали цветам встроенных областей задач. Эти цвета будут динамически обновляться при изменении темы интерфейса Office.</span><span class="sxs-lookup"><span data-stu-id="935ef-p103">**In task pane add-ins for PowerPoint**. Use the Office UI theme classes of OfficeThemes.css to specify the same fonts and background colors used in the UI so that your task pane add-ins will match the colors of built-in task panes - and those colors will dynamically update if a user changes the Office UI theme.</span></span>

### <a name="document-theme-colors"></a><span data-ttu-id="935ef-111">Цвета темы документа</span><span class="sxs-lookup"><span data-stu-id="935ef-111">Document theme colors</span></span>

<span data-ttu-id="935ef-p104">Каждая тема документа Office определяет 12 цветов. Десять из них доступны при выборе шрифта, фона и других цветовых настроек презентации с помощью палитры.</span><span class="sxs-lookup"><span data-stu-id="935ef-p104">Every Office document theme defines 12 colors. Ten of these colors are available when you set font, background, and other color settings in a presentation with the color picker.</span></span>

![Цветовая палитра](../images/office15-app-color-palette.png)

<span data-ttu-id="935ef-115">Чтобы просмотреть или настроить полный набор из 12 цветов темы в PowerPoint, в группе **варианты** на вкладке **конструктор** нажмите кнопку **Дополнительно** , а затем выберите пункт **цвета** > **Настройка цветов** , чтобы открыть диалоговое окно " **Создание новых цветов темы** ".</span><span class="sxs-lookup"><span data-stu-id="935ef-115">To view or customize the full set of 12 theme colors in PowerPoint, in the **Variants** group on the **Design** tab, click the **More** drop-down - then select **Colors** > **Customize Colors** to display the **Create New Theme Colors** dialog box.</span></span>

![Создание диалогового окна с цветами темы](../images/office15-app-create-new-theme-colors.png)

<span data-ttu-id="935ef-p105">Первые четыре цвета предназначены для текста и фона. Текст, выполненный в светлых тонах, всегда лучше читается на темном фоне, а текст темных тонов — на светлом фоне. Следующие шесть цветов — это контрастные цвета, которые всегда четко видны на четырех возможных фоновых цветах. Последние два цвета применяются для непросмотренных и просмотренных гиперссылок.</span><span class="sxs-lookup"><span data-stu-id="935ef-p105">The first four colors are for text and backgrounds. Text that is created with the light colors will always be legible over the dark colors, and text that is created with dark colors will always be legible over the light colors. The next six are accent colors that are always visible over the four potential background colors. The last two colors are for hyperlinks and followed hyperlinks.</span></span>

### <a name="document-theme-fonts"></a><span data-ttu-id="935ef-121">Шрифты темы документа</span><span class="sxs-lookup"><span data-stu-id="935ef-121">Document theme fonts</span></span>

<span data-ttu-id="935ef-122">В каждой теме документа Office определено два шрифта: один для заголовков и другой для основного текста.</span><span class="sxs-lookup"><span data-stu-id="935ef-122">Every Office document theme also defines two fonts -- one for headings and one for body text.</span></span> <span data-ttu-id="935ef-123">PowerPoint использует их для создания автоматических текстовых стилей.</span><span class="sxs-lookup"><span data-stu-id="935ef-123">PowerPoint uses these fonts to construct automatic text styles.</span></span> <span data-ttu-id="935ef-124">Кроме того, они используются в коллекциях текстовых **экспресс-стилей** и **WordArt**.</span><span class="sxs-lookup"><span data-stu-id="935ef-124">In addition, **Quick Styles** galleries for text and **WordArt** use these same theme fonts.</span></span> <span data-ttu-id="935ef-125">Эти два шрифта отображаются вверху средства выбора шрифтов.</span><span class="sxs-lookup"><span data-stu-id="935ef-125">These two fonts are available as the first two selections when you select fonts with the font picker.</span></span>

![Средство выбора шрифтов](../images/office15-app-font-picker.png)

<span data-ttu-id="935ef-127">Чтобы просмотреть или настроить шрифты темы в PowerPoint, в группе **варианты** на вкладке **конструктор** нажмите кнопку **раскрывающегося** списка, а затем выберите пункт **шрифты** > **Настройка шрифтов** , чтобы открыть диалоговое окно **Создание новых шрифтов темы** .</span><span class="sxs-lookup"><span data-stu-id="935ef-127">To view or customize theme fonts in PowerPoint, in the **Variants** group on the **Design** tab, click the **More** drop-down - then select **Fonts** > **Customize Fonts** to display the **Create New Theme Fonts** dialog box.</span></span>

![Создание диалогового окна со шрифтами темы](../images/office15-app-create-new-theme-fonts.png)

### <a name="office-ui-theme-fonts-and-colors"></a><span data-ttu-id="935ef-129">Шрифты и цвета темы для пользовательского интерфейса Office</span><span class="sxs-lookup"><span data-stu-id="935ef-129">Office UI theme fonts and colors</span></span>

<span data-ttu-id="935ef-130">Office также позволяет выбирать между несколькими стандартными темами, которые определяют несколько цветов и шрифтов, используемых в пользовательском интерфейсе всех приложений Office.</span><span class="sxs-lookup"><span data-stu-id="935ef-130">Office also lets you choose between several predefined themes that specify some of the colors and fonts used in the UI of all Office applications.</span></span> <span data-ttu-id="935ef-131">Для этого используется раскрывающийся список "**тема Office** " для**учетной записи** >  **файла** > (из любого приложения Office).</span><span class="sxs-lookup"><span data-stu-id="935ef-131">To do that, you use the **File** > **Account** > **Office Theme** drop-down (from any Office application).</span></span>

![Раскрывающийся список тем Office](../images/office15-app-office-theme-picker.png)

<span data-ttu-id="935ef-p108">Файл OfficeThemes.css содержит классы, которые можно использовать в надстройках области задач для PowerPoint, чтобы в них применялись аналогичные шрифты и цвета. Это позволит вам создавать свои надстройки области задач, внешний вид которых совпадает с внешним видом встроенных областей задач.</span><span class="sxs-lookup"><span data-stu-id="935ef-p108">OfficeThemes.css includes classes that you can use in your task pane add-ins for PowerPoint so they will use these same fonts and colors. This lets you design your task pane add-ins that match the appearance of built-in task panes.</span></span>

## <a name="using-officethemescss"></a><span data-ttu-id="935ef-135">Использование OfficeThemes.css</span><span class="sxs-lookup"><span data-stu-id="935ef-135">Using OfficeThemes.css</span></span>

<span data-ttu-id="935ef-p109">Использование файла OfficeThemes.css вместе с контентными надстройками для PowerPoint позволит вам согласовать внешний вид надстройка с темой презентации, а использование этого файла с надстройками областей задач для PowerPoint позволит согласовать внешний вид надстройка со шрифтами и цветами пользовательского интерфейса Office.</span><span class="sxs-lookup"><span data-stu-id="935ef-p109">Using the OfficeThemes.css file with your content add-ins for PowerPoint lets you coordinate the appearance of your add-in with the theme applied to the presentation it's running with. Using the OfficeThemes.css file with your task pane add-ins for PowerPoint lets you coordinate the appearance of your add-in with the fonts and colors of the Office UI.</span></span>

### <a name="adding-the-officethemescss-file-to-your-project"></a><span data-ttu-id="935ef-138">Добавление файла OfficeThemes.css в проект</span><span class="sxs-lookup"><span data-stu-id="935ef-138">Adding the OfficeThemes.css file to your project</span></span>

<span data-ttu-id="935ef-139">Чтобы добавить файл OfficeThemes.css и ссылку на него в проекте надстройка, выполните описанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="935ef-139">Use the following steps to add and reference the OfficeThemes.css file to your add-in project.</span></span>

#### <a name="to-add-officethemescss-to-your-visual-studio-project"></a><span data-ttu-id="935ef-140">Добавление файла OfficeThemes.css в проект Visual Studio</span><span class="sxs-lookup"><span data-stu-id="935ef-140">To add OfficeThemes.css to your Visual Studio project</span></span>

> [!NOTE]
> <span data-ttu-id="935ef-141">Действия этой процедуры применимы только к Visual Studio 2015.</span><span class="sxs-lookup"><span data-stu-id="935ef-141">The steps in this procedure only apply to Visual Studio 2015.</span></span> <span data-ttu-id="935ef-142">Если вы используете Visual Studio 2019, файл OfficeThemes. CSS создается автоматически для всех создаваемых проектов надстроек PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="935ef-142">If you are using Visual Studio 2019, the OfficeThemes.css file is created automatically for any new PowerPoint add-in projects that you create.</span></span>

1. <span data-ttu-id="935ef-143">В **обозревателе решений** щелкните правой кнопкой мыши папку **Content** в проекте _\*\*имя_проекта**_**Web\*\* и выберите **Добавить** > **Таблица стилей**.</span><span class="sxs-lookup"><span data-stu-id="935ef-143">In **Solution Explorer**, right-click the **Content** folder in the _**project_name**_**Web** project, choose **Add**, and then select **Style Sheet**.</span></span>
    
2. <span data-ttu-id="935ef-144">Назовите новую таблицу стилей **OfficeThemes**.</span><span class="sxs-lookup"><span data-stu-id="935ef-144">Name the new style sheet **OfficeThemes**.</span></span>
    
   > [!IMPORTANT]
   > <span data-ttu-id="935ef-145">Таблица стилей должна называться OfficeThemes, в противном случае не будет работать динамическое обновление шрифтов и цветов надстроек.</span><span class="sxs-lookup"><span data-stu-id="935ef-145">The style sheet must be named OfficeThemes, or the feature that dynamically updates add-in fonts and colors when a user changes the theme won't work.</span></span>
   
3. <span data-ttu-id="935ef-146">Удалите из файла класс **body** по умолчанию (`body {}`) и скопируйте в файл представленный ниже CSS-код.</span><span class="sxs-lookup"><span data-stu-id="935ef-146">Delete the default **body** class (`body {}`) in the file, and copy and paste the following CSS code into the file.</span></span>
    
    ```css
    /* The following classes describe the common theme information for office documents */ 

    /* Basic Font and Background Colors for text */ 
    .office-docTheme-primary-fontColor { color:#000000; } 
    .office-docTheme-primary-bgColor { background-color:#ffffff; } 
    .office-docTheme-secondary-fontColor { color: #000000; } 
    .office-docTheme-secondary-bgColor { background-color: #ffffff; } 

    /* Accent color definitions for fonts */ 
    .office-contentAccent1-color { color:#5b9bd5; } 
    .office-contentAccent2-color { color:#ed7d31; } 
    .office-contentAccent3-color { color:#a5a5a5; } 
    .office-contentAccent4-color { color:#ffc000; } 
    .office-contentAccent5-color { color:#4472c4; } 
    .office-contentAccent6-color { color:#70ad47; } 

    /* Accent color for backgrounds */ 
    .office-contentAccent1-bgColor { background-color:#5b9bd5; } 
    .office-contentAccent2-bgColor { background-color:#ed7d31; } 
    .office-contentAccent3-bgColor { background-color:#a5a5a5; } 
    .office-contentAccent4-bgColor { background-color:#ffc000; } 
    .office-contentAccent5-bgColor { background-color:#4472c4; } 
    .office-contentAccent6-bgColor { background-color:#70ad47; } 

    /* Accent color for borders */ 
    .office-contentAccent1-borderColor { border-color:#5b9bd5; } 
    .office-contentAccent2-borderColor { border-color:#ed7d31; } 
    .office-contentAccent3-borderColor { border-color:#a5a5a5; } 
    .office-contentAccent4-borderColor { border-color:#ffc000; } 
    .office-contentAccent5-borderColor { border-color:#4472c4; } 
    .office-contentAccent6-borderColor { border-color:#70ad47; } 

    /* links */ 
    .office-a { color: #0563c1; } 
    .office-a:visited { color: #954f72; } 

    /* Body Fonts */ 
    .office-bodyFont-eastAsian { } /* East Asian name of the Font */ 
    .office-bodyFont-latin { font-family:"Calibri"; } /* Latin name of the Font */ 
    .office-bodyFont-script { } /* Script name of the Font */ 
    .office-bodyFont-localized { font-family:"Calibri"; } /* Localized name of the Font. Corresponds to the default font of the culture currently used in Office.*/ 

    /* Headers Font */ 
    .office-headerFont-eastAsian { } 
    .office-headerFont-latin { font-family:"Calibri Light"; } 
    .office-headerFont-script { } 
    .office-headerFont-localized { font-family:"Calibri Light"; } 

    /* The following classes define font and background colors for Office UI themes. These classes should only be used in task pane add-ins */ 

    /* Basic Font and Background Colors for PPT */ 
    .office-officeTheme-primary-fontColor { color:#b83b1d; } 
    .office-officeTheme-primary-bgColor { background-color:#dedede; } 
    .office-officeTheme-secondary-fontColor { color:#262626; } 
    .office-officeTheme-secondary-bgColor { background-color:#ffffff; }
    ```
4. <span data-ttu-id="935ef-147">Если вы используете отличный от Visual Studio инструмент для создания надстройка, скопируйте CSS-код из третьего шага в текстовый файл, сохранив его под именем OfficeThemes.css.</span><span class="sxs-lookup"><span data-stu-id="935ef-147">If you are using a tool other than Visual Studio to create your add-in, copy the CSS code from step 3 into a text file, making sure to save the file as OfficeThemes.css.</span></span>   

### <a name="referencing-officethemescss-in-your-add-ins-html-pages"></a><span data-ttu-id="935ef-148">Добавление ссылок на файл OfficeThemes.css в HTML-страницах надстройки</span><span class="sxs-lookup"><span data-stu-id="935ef-148">Referencing OfficeThemes.css in your add-in's HTML pages</span></span>

<span data-ttu-id="935ef-149">Чтобы использовать файл OfficeThemes.css в проекте надстройки, добавьте тег `<link>`, который ссылается на файл OfficeThemes.css, внутри тега `<head>` веб-страницы надстройки (HTML, ASPX или PHP) в следующем формате:</span><span class="sxs-lookup"><span data-stu-id="935ef-149">To use the OfficeThemes.css file in your add-in project, add a `<link>` tag that references the OfficeThemes.css file inside the `<head>` tag of the web pages (such as an .html, .aspx, or .php file) that implement the UI of your add-in in this format:</span></span>

```HTML
<link href="<local_path_to_OfficeThemes.css>" rel="stylesheet" type="text/css" />
```

<span data-ttu-id="935ef-150">Чтобы сделать это в Visual Studio, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="935ef-150">To do this in Visual Studio, follow these steps.</span></span>

#### <a name="to-reference-officethemescss-in-your-add-in-for-powerpoint"></a><span data-ttu-id="935ef-151">Добавление ссылки на OfficeThemes.css в надстройке для PowerPoint</span><span class="sxs-lookup"><span data-stu-id="935ef-151">To reference OfficeThemes.css in your add-in for PowerPoint</span></span>

1. <span data-ttu-id="935ef-152">Выберите **Создание нового проекта**.</span><span class="sxs-lookup"><span data-stu-id="935ef-152">Choose **Create a new project**.</span></span>

2. <span data-ttu-id="935ef-153">Используя поле поиска, введите **надстройка**.</span><span class="sxs-lookup"><span data-stu-id="935ef-153">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="935ef-154">Выберите вариант **Веб-надстройка PowerPoint** и нажмите кнопку **Далее**.</span><span class="sxs-lookup"><span data-stu-id="935ef-154">Choose **PowerPoint Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="935ef-155">Присвойте проекту имя и нажмите кнопку **Создать**.</span><span class="sxs-lookup"><span data-stu-id="935ef-155">Name your project and select **Create**.</span></span>

3. <span data-ttu-id="935ef-156">В диалоговом окне **Создание надстройки Office** выберите **Добавить новые функции в PowerPoint**, а затем нажмите кнопку **Готово**, чтобы создать проект.</span><span class="sxs-lookup"><span data-stu-id="935ef-156">In the **Create Office Add-in** dialog window, choose **Add new functionalities to PowerPoint**, and then choose **Finish** to create the project.</span></span>

4. <span data-ttu-id="935ef-p112">Visual Studio создаст решение, и в **обозревателе решений** появятся два соответствующих проекта. В Visual Studio откроется файл **Home.html**.</span><span class="sxs-lookup"><span data-stu-id="935ef-p112">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

5. <span data-ttu-id="935ef-159">На HTML-страницах надстройки, например Home.html в шаблоне по умолчанию, добавьте следующий тег `<link>`, который ссылается на файл OfficeThemes.css, внутри тега `<head>`:</span><span class="sxs-lookup"><span data-stu-id="935ef-159">In the HTML pages that implement the UI of your add-in, such as Home.html in the default template, add the following `<link>` tag inside the `<head>` tag that references the OfficeThemes.css file:</span></span>
    
    ```HTML
    <link href="../../Content/OfficeThemes.css" rel="stylesheet" type="text/css" />
    ```

<span data-ttu-id="935ef-160">При использовании другого инструмента добавьте тег `<link>` в таком же формате, указав относительный путь к копии файла OfficeThemes.css, которая будет разворачиваться с вашей надстройкой.</span><span class="sxs-lookup"><span data-stu-id="935ef-160">If you are creating your add-in with a tool other than Visual Studio, add a `<link>` tag with the same format specifying a relative path to the copy of OfficeThemes.css that will be deployed with your add-in.</span></span>

### <a name="using-officethemescss-document-theme-classes-in-your-content-add-ins-html-page"></a><span data-ttu-id="935ef-161">Использование классов OfficeThemes.css для темы документа на HTML-странице контентной надстройки</span><span class="sxs-lookup"><span data-stu-id="935ef-161">Using OfficeThemes.css document theme classes in your content add-in's HTML page</span></span>

<span data-ttu-id="935ef-p113">Ниже представлен простой пример HTML-кода в контентной надстройке, которая использует классы OfficeTheme.css для темы документа. Более подробные сведения о классах OfficeThemes.css, которые соответствуют используемым в теме документа 12 цветам и 2 шрифтам, можно узнать в разделе [Классы тем для контентных надстроек](#theme-classes-for-content-add-ins).</span><span class="sxs-lookup"><span data-stu-id="935ef-p113">The following shows a simple example of HTML in a content add-in that uses the OfficeTheme.css document theme classes. For details about the OfficeThemes.css classes that correspond to the 12 colors and 2 fonts used in a document theme, see [Theme classes for content add-ins](#theme-classes-for-content-add-ins).</span></span>

```HTML
<body>
    <div id="themeSample" class="office-docTheme-primary-fontColor ">
        <h1 class="office-headerFont-latin">Hello world!</h1> 
        <h1 class="office-headerFont-latin office-contentAccent1-bgColor">Hello world!</h1> 
        <h1 class="office-headerFont-latin office-contentAccent2-bgColor">Hello world!</h1> 
        <h1 class="office-headerFont-latin office-contentAccent3-bgColor">Hello world!</h1> 
        <h1 class="office-headerFont-latin office-contentAccent4-bgColor">Hello world!</h1> 
        <h1 class="office-headerFont-latin office-contentAccent5-bgColor">Hello world!</h1> 
        <h1 class="office-headerFont-latin office-contentAccent6-bgColor">Hello world!</h1> 
        <p class="office-bodyFont-latin office-docTheme-secondary-fontColor">Hello world!</p> 
    </div>
</body>
```

<span data-ttu-id="935ef-164">В среде выполнения при вставке в презентацию, которая использует **тему Office**по умолчанию, контентная Надстройка отображается подобным образом.</span><span class="sxs-lookup"><span data-stu-id="935ef-164">At runtime, when inserted into a presentation that uses the default **Office Theme**, the content add-in is rendered like this.</span></span>

![Контентное приложение, в котором используется тема Office](../images/office15-app-content-app-office-theme.png)

<span data-ttu-id="935ef-p114">Если вы измените тему презентации или настроите текущую тему, шрифты и цвета, указанные с помощью классов OfficeThemes.css, динамически обновятся. Если презентация, в которую вставляется надстройка, использует тему **Аспект**, описанная выше HTML-страница надстройки будет выглядеть так:</span><span class="sxs-lookup"><span data-stu-id="935ef-p114">If you change the presentation to use another theme or customize the presentation's theme, the fonts and colors specified with OfficeThemes.css classes will dynamically update to correspond to the fonts and colors of the presentation's theme. Using the same HTML example as above, if the presentation the add-in is inserted into uses the **Facet** theme, the add-in rendering will look like this.</span></span>

![Контентное приложение, в котором используется тема "Аспект"](../images/office15-app-content-app-facet-theme.png)


### <a name="using-officethemescss-office-ui-theme-classes-in-your-task-pane-add-ins-html-page"></a><span data-ttu-id="935ef-169">Использование классов OfficeThemes.css для темы пользовательского интерфейса Office в HTML-странице надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="935ef-169">Using OfficeThemes.css Office UI theme classes in your task pane add-in's HTML page</span></span>

<span data-ttu-id="935ef-170">Помимо темы документа, пользователи могут настраивать цветовую схему приложения Office. Для этого используется раскрывающийся список **Файл** > **Учетная запись** > **Тема Office**.</span><span class="sxs-lookup"><span data-stu-id="935ef-170">In addition to the document theme, users can customize the color scheme of the Office user interface for all Office applications using the **File** > **Account** > **Office Theme** drop-down box.</span></span>

<span data-ttu-id="935ef-p115">Ниже показан пример простого HTML-кода в надстройка области задач, который использует классы OfficeTheme.css для указания цвета шрифта и фона. Более подробную информацию о классах OfficeThemes.css, которые соответствуют шрифтам и цветам темы пользовательского интерфейса Office, можно получить в разделе [Классы тем для надстроек области задач](#theme-classes-for-task-pane-add-ins).</span><span class="sxs-lookup"><span data-stu-id="935ef-p115">The following shows a simple example of HTML in a task pane add-in that uses OfficeTheme.css classes to specify font color and background color. For details about the OfficeThemes.css classes that correspond to fonts and colors of the Office UI theme, see [Theme classes for task pane add-ins](#theme-classes-for-task-pane-add-ins).</span></span>

```HTML
<body> 
    <div id="content-header" class="office-officeTheme-primary-fontColor office-officeTheme-primary-bgColor"> 
        <div class="padding">
            <h1>Welcome</h1>
        </div> 
    </div> 
    <div id="content-main" class="office-officeTheme-secondary-fontColor office-officeTheme-secondary-bgColor"> 
        <div class="padding"> 
            <p>Add home screen content here.</p> 
            <p>For example:</p> 
            <button id="get-data-from-selection">Get data from selection</button> 
            <p><a target="_blank" class="office-a" href="https://go.microsoft.com/fwlink/?LinkId=276812">Find more samples online...</a></p>
        </div>
    </div>
</body> 
```

<br/>

<span data-ttu-id="935ef-173">Если в PowerPoint выбрать в раскрывающемся списке **Файл** > **Учетная запись** > **Тема Office** значение **Белая**, то надстройка области задач будет выглядеть так:</span><span class="sxs-lookup"><span data-stu-id="935ef-173">When running in PowerPoint with **File** > **Account** > **Office Theme** set to **White**, the task pane add-in is rendered like this.</span></span>

![Область задач с белой темой Office](../images/office15-app-task-pane-theme-white.png)

<br/>

<span data-ttu-id="935ef-175">Если вы измените **тему Office** на **темно-серую**, шрифты и цвета, указанные с помощью классов в OfficeThemes.css, динамически обновятся и станут выглядеть так:</span><span class="sxs-lookup"><span data-stu-id="935ef-175">If you change **OfficeTheme** to **Dark Gray**, the fonts and colors specified with OfficeThemes.css classes will dynamically update to render like this.</span></span>

![Область задач с темно-серой темой Office](../images/office15-app-task-pane-theme-dark-gray.png)

<br/>

## <a name="officethemecss-classes"></a><span data-ttu-id="935ef-177">Классы OfficeTheme.css</span><span class="sxs-lookup"><span data-stu-id="935ef-177">OfficeTheme.css classes</span></span>

<span data-ttu-id="935ef-178">Файл OfficeThemes.css содержит два набора классов, которые вы можете использовать с контентными надстройками и надстройками области задач для PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="935ef-178">The OfficeThemes.css file contains two sets of classes you can use with your content and task pane add-ins for PowerPoint.</span></span>

### <a name="theme-classes-for-content-add-ins"></a><span data-ttu-id="935ef-179">Классы тем для контентных надстроек</span><span class="sxs-lookup"><span data-stu-id="935ef-179">Theme classes for content add-ins</span></span>

<span data-ttu-id="935ef-p116">Файл OfficeThemes.css предоставляет классы, соответствующие 2 шрифтам и 12 цветам, используемым в теме документа. Эти классы предназначены для использования с контентными надстройками для PowerPoint, чтобы шрифты и цвета вашей надстройки были согласованы с презентацией, в которую она вставлена.</span><span class="sxs-lookup"><span data-stu-id="935ef-p116">The OfficeThemes.css file provides classes that correspond to the 2 fonts and 12 colors used in a document theme. These classes are appropriate to use with content add-ins for PowerPoint so that your add-in's fonts and colors will be coordinated with the presentation it's inserted into.</span></span>

#### <a name="theme-fonts-for-content-add-ins"></a><span data-ttu-id="935ef-182">Шрифты тем для контентных надстроек</span><span class="sxs-lookup"><span data-stu-id="935ef-182">Theme fonts for content add-ins</span></span>

|<span data-ttu-id="935ef-183">**Класс**</span><span class="sxs-lookup"><span data-stu-id="935ef-183">**Class**</span></span>|<span data-ttu-id="935ef-184">**Описание**</span><span class="sxs-lookup"><span data-stu-id="935ef-184">**Description**</span></span>|
|:-----|:-----|
| `office-bodyFont-eastAsian`|<span data-ttu-id="935ef-185">Восточноазиатское имя шрифта основного текста.</span><span class="sxs-lookup"><span data-stu-id="935ef-185">East Asian name of the body font.</span></span>|
| `office-bodyFont-latin`|<span data-ttu-id="935ef-p117">Латинское название шрифта основного текста. По умолчанию "Calabri"</span><span class="sxs-lookup"><span data-stu-id="935ef-p117">Latin name of the body font. Default "Calabri"</span></span>|
| `office-bodyFont-script`|<span data-ttu-id="935ef-188">Имя сценария шрифта основного текста.</span><span class="sxs-lookup"><span data-stu-id="935ef-188">Script name of the body font.</span></span>|
| `office-bodyFont-localized`|<span data-ttu-id="935ef-p118">Локализованное имя шрифта основного текста. Задает название шрифта по умолчанию в соответствии с текущей культурой, используемой в Office</span><span class="sxs-lookup"><span data-stu-id="935ef-p118">Localized name of the body font. Specifies the default font name according to the culture currently used in Office.</span></span>|
| `office-headerFont-eastAsian`|<span data-ttu-id="935ef-191">Восточноазиатское название шрифта заголовков</span><span class="sxs-lookup"><span data-stu-id="935ef-191">East Asian name of the headers font.</span></span>|
| `office-headerFont-latin`|<span data-ttu-id="935ef-p119">Латинское название шрифта заголовков. По умолчанию "Calabri Light"</span><span class="sxs-lookup"><span data-stu-id="935ef-p119">Latin name of the headers font. Default "Calabri Light"</span></span>|
| `office-headerFont-script`|<span data-ttu-id="935ef-194">Имя сценариев шрифта заголовков</span><span class="sxs-lookup"><span data-stu-id="935ef-194">Script name of the headers font.</span></span>|
| `office-headerFont-localized`|<span data-ttu-id="935ef-p120">Локализованное название шрифта заголовков. Задает название шрифта по умолчанию в соответствии с текущей культурой, используемой в Office</span><span class="sxs-lookup"><span data-stu-id="935ef-p120">Localized name of the headers font. Specifies the default font name according to the culture currently used in Office.</span></span>|

<br/>

#### <a name="theme-colors-for-content-add-ins"></a><span data-ttu-id="935ef-197">Цвета тем для контентных надстроек</span><span class="sxs-lookup"><span data-stu-id="935ef-197">Theme colors for content add-ins</span></span>

|<span data-ttu-id="935ef-198">**Класс**</span><span class="sxs-lookup"><span data-stu-id="935ef-198">**Class**</span></span>|<span data-ttu-id="935ef-199">**Описание**</span><span class="sxs-lookup"><span data-stu-id="935ef-199">**Description**</span></span>|
|:-----|:-----|
| `office-docTheme-primary-fontColor`|<span data-ttu-id="935ef-p121">Основной цвет шрифта. По умолчанию #000000</span><span class="sxs-lookup"><span data-stu-id="935ef-p121">Primary font color. Default #000000</span></span>|
| `office-docTheme-primary-bgColor`|<span data-ttu-id="935ef-p122">Основной цвет фона шрифта. По умолчанию #FFFFFF</span><span class="sxs-lookup"><span data-stu-id="935ef-p122">Primary font background color. Default #FFFFFF</span></span>|
| `office-docTheme-secondary-fontColor`|<span data-ttu-id="935ef-p123">Дополнительный цвет шрифта. По умолчанию #000000</span><span class="sxs-lookup"><span data-stu-id="935ef-p123">Secondary font color. Default #000000</span></span>|
| `office-docTheme-secondary-bgColor`|<span data-ttu-id="935ef-p124">Дополнительный цвет фона шрифта. По умолчанию #FFFFFF</span><span class="sxs-lookup"><span data-stu-id="935ef-p124">Secondary font background color. Default #FFFFFF</span></span>|
| `office-contentAccent1-color`|<span data-ttu-id="935ef-p125">Контрастный цвет шрифта 1. По умолчанию #5B9BD5</span><span class="sxs-lookup"><span data-stu-id="935ef-p125">Font accent color 1. Default #5B9BD5</span></span>|
| `office-contentAccent2-color`|<span data-ttu-id="935ef-p126">Контрастный цвет шрифта 2. По умолчанию #ED7D31</span><span class="sxs-lookup"><span data-stu-id="935ef-p126">Font accent color 2. Default #ED7D31</span></span>|
| `office-contentAccent3-color`|<span data-ttu-id="935ef-p127">Контрастный цвет шрифта 3. По умолчанию #A5A5A5</span><span class="sxs-lookup"><span data-stu-id="935ef-p127">Font accent color 3. Default #A5A5A5</span></span>|
| `office-contentAccent4-color`|<span data-ttu-id="935ef-p128">Контрастный цвет шрифта 4. По умолчанию #FFC000</span><span class="sxs-lookup"><span data-stu-id="935ef-p128">Font accent color 4. Default #FFC000</span></span>|
| `office-contentAccent5-color`|<span data-ttu-id="935ef-p129">Контрастный цвет шрифта 5. По умолчанию #4472C4</span><span class="sxs-lookup"><span data-stu-id="935ef-p129">Font accent color 5. Default #4472C4</span></span>|
| `office-contentAccent6-color`|<span data-ttu-id="935ef-p130">Контрастный цвет шрифта 6. По умолчанию #70AD47</span><span class="sxs-lookup"><span data-stu-id="935ef-p130">Font accent color 6. Default #70AD47</span></span>|
| `office-contentAccent1-bgColor`|<span data-ttu-id="935ef-p131">Контрастный цвет фона 1. По умолчанию #5B9BD5</span><span class="sxs-lookup"><span data-stu-id="935ef-p131">Background accent color 1. Default #5B9BD5</span></span>|
| `office-contentAccent2-bgColor`|<span data-ttu-id="935ef-p132">Контрастный цвет фона 2. По умолчанию #ED7D31</span><span class="sxs-lookup"><span data-stu-id="935ef-p132">Background accent color 2. Default #ED7D31</span></span>|
| `office-contentAccent3-bgColor`|<span data-ttu-id="935ef-p133">Контрастный цвет фона 3. По умолчанию #A5A5A5</span><span class="sxs-lookup"><span data-stu-id="935ef-p133">Background accent color 3. Default #A5A5A5</span></span>|
| `office-contentAccent4-bgColor`|<span data-ttu-id="935ef-p134">Контрастный цвет фона 4. По умолчанию #FFC000</span><span class="sxs-lookup"><span data-stu-id="935ef-p134">Background accent color 4. Default #FFC000</span></span>|
| `office-contentAccent5-bgColor`|<span data-ttu-id="935ef-p135">Контрастный цвет фона 5. По умолчанию #4472C4</span><span class="sxs-lookup"><span data-stu-id="935ef-p135">Background accent color 5. Default #4472C4</span></span>|
| `office-contentAccent6-bgColor`|<span data-ttu-id="935ef-p136">Контрастный цвет фона 6. По умолчанию #70AD47</span><span class="sxs-lookup"><span data-stu-id="935ef-p136">Background accent color 6. Default #70AD47</span></span>|
| `office-contentAccent1-borderColor`|<span data-ttu-id="935ef-p137">Контрастный цвет границы 1. По умолчанию #5B9BD5</span><span class="sxs-lookup"><span data-stu-id="935ef-p137">Border accent color 1. Default #5B9BD5</span></span>|
| `office-contentAccent2-borderColor`|<span data-ttu-id="935ef-p138">Контрастный цвет границы 2. По умолчанию #ED7D31</span><span class="sxs-lookup"><span data-stu-id="935ef-p138">Border accent color 2. Default #ED7D31</span></span>|
| `office-contentAccent3-borderColor`|<span data-ttu-id="935ef-p139">Контрастный цвет границы 3. По умолчанию #A5A5A5</span><span class="sxs-lookup"><span data-stu-id="935ef-p139">Border accent color 3. Default #A5A5A5</span></span>|
| `office-contentAccent4-borderColor`|<span data-ttu-id="935ef-p140">Контрастный цвет границы 4. По умолчанию #FFC000</span><span class="sxs-lookup"><span data-stu-id="935ef-p140">Border accent color 4. Default #FFC000</span></span>|
| `office-contentAccent5-borderColor`|<span data-ttu-id="935ef-p141">Контрастный цвет границы 5. По умолчанию #4472C4</span><span class="sxs-lookup"><span data-stu-id="935ef-p141">Border accent color 5. Default #4472C4</span></span>|
| `office-contentAccent6-borderColor`|<span data-ttu-id="935ef-p142">Контрастный цвет границы 6. По умолчанию #70AD47</span><span class="sxs-lookup"><span data-stu-id="935ef-p142">Border accent color 6. Default #70AD47</span></span>|
| `office-a`|<span data-ttu-id="935ef-p143">Цвет гиперссылки. По умолчанию #0563C1</span><span class="sxs-lookup"><span data-stu-id="935ef-p143">Hyperlink color. Default #0563C1</span></span>|
| `office-a:visited`|<span data-ttu-id="935ef-p144">Цвет просмотренной гиперссылки. По умолчанию #954F72</span><span class="sxs-lookup"><span data-stu-id="935ef-p144">Followed hyperlink color. Default #954F72</span></span>|

<br/>

<span data-ttu-id="935ef-248">На следующем снимке экрана представлены примеры всех классов цветов темы (за исключением двух цветов для гиперссылок), указанных для текста надстройка при использовании темы Office по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="935ef-248">The following screenshot shows examples of all of the theme color classes (except for the two hyperlink colors) assigned to add-in text when using the default Office theme.</span></span>

![Пример цветов темы Office по умолчанию](../images/office15-app-default-office-theme-colors.png)


### <a name="theme-classes-for-task-pane-add-ins"></a><span data-ttu-id="935ef-250">Классы тем для надстроек области задач</span><span class="sxs-lookup"><span data-stu-id="935ef-250">Theme classes for task pane add-ins</span></span>

<span data-ttu-id="935ef-p145">В файле OfficeThemes.css представлены классы, соответствующие 4 цветам, которые указаны для шрифтов и фонов, использующихся темой пользовательского интерфейса приложения Office. Эти классы предназначены для использования с надстройками области задач для PowerPoint, поэтому цвета вашей надстройки будут согласованы с цветами других встроенных областей задач в Office.</span><span class="sxs-lookup"><span data-stu-id="935ef-p145">The OfficeThemes.css file provides classes that correspond to the 4 colors assigned to fonts and backgrounds used by the Office application UI theme. These classes are appropriate to use with task add-ins for PowerPoint so that your add-in's colors will be coordinated with the other built-in task panes in Office.</span></span>

#### <a name="theme-font-and-background-colors-for-task-pane-add-ins"></a><span data-ttu-id="935ef-253">Цвета шрифта и фона тем для надстроек области задач</span><span class="sxs-lookup"><span data-stu-id="935ef-253">Theme font and background colors for task pane add-ins</span></span>

|<span data-ttu-id="935ef-254">**Класс**</span><span class="sxs-lookup"><span data-stu-id="935ef-254">**Class**</span></span>|<span data-ttu-id="935ef-255">**Описание**</span><span class="sxs-lookup"><span data-stu-id="935ef-255">**Description**</span></span>|
|:-----|:-----|
| `office-officeTheme-primary-fontColor`|<span data-ttu-id="935ef-p146">Основной цвет шрифта. Значение по умолчанию — #B83B1D.</span><span class="sxs-lookup"><span data-stu-id="935ef-p146">Primary font color. Default #B83B1D</span></span>|
| `office-officeTheme-primary-bgColor`|<span data-ttu-id="935ef-p147">Основной цвет фона. Значение по умолчанию — #DEDEDE.</span><span class="sxs-lookup"><span data-stu-id="935ef-p147">Primary background color. Default #DEDEDE</span></span>|
| `office-officeTheme-secondary-fontColor`|<span data-ttu-id="935ef-p148">Дополнительный цвет шрифта. По умолчанию #262626</span><span class="sxs-lookup"><span data-stu-id="935ef-p148">Secondary font color. Default #262626</span></span>|
| `office-officeTheme-secondary-bgColor`|<span data-ttu-id="935ef-p149">Дополнительный цвет фона. Значение по умолчанию — #FFFFFF.</span><span class="sxs-lookup"><span data-stu-id="935ef-p149">Secondary background color. Default #FFFFFF</span></span>|

## <a name="see-also"></a><span data-ttu-id="935ef-264">См. также</span><span class="sxs-lookup"><span data-stu-id="935ef-264">See also</span></span>

- [<span data-ttu-id="935ef-265">Создание контентных надстроек и надстроек области задач для PowerPoint</span><span class="sxs-lookup"><span data-stu-id="935ef-265">Create content and task pane add-ins for PowerPoint</span></span>](../powerpoint/powerpoint-add-ins.md)
