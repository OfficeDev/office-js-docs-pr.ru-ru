---
title: Проверка манифеста надстройки Office
description: Узнайте, как проверить манифест надстройки Office с помощью схемы XML и других средств.
ms.date: 12/31/2019
localization_priority: Priority
ms.openlocfilehash: 09b5841a0180d8cb730ec8b479df1386a0749b60
ms.sourcegitcommit: d5ac9284d1e96dc91a9168d7641e44d88535e1a7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/31/2019
ms.locfileid: "40914904"
---
# <a name="validate-an-office-add-ins-manifest"></a><span data-ttu-id="fc1d6-103">Проверка манифеста надстройки Office</span><span class="sxs-lookup"><span data-stu-id="fc1d6-103">Validate an Office Add-in manifest</span></span>

<span data-ttu-id="fc1d6-104">Может потребоваться проверить файл манифеста надстройки, чтобы убедиться в его правильности и полноте.</span><span class="sxs-lookup"><span data-stu-id="fc1d6-104">You may want to validate your add-in's manifest file to ensure that it's correct and complete.</span></span> <span data-ttu-id="fc1d6-105">Проверка может также выявлять проблемы, которые приводят к появлению ошибки "Манифест надстройки недействителен" при попытке загрузить неопубликованную надстройку.</span><span class="sxs-lookup"><span data-stu-id="fc1d6-105">Validation can also identify issues that are causing the error "Your add-in manifest is not valid" when you attempt to sideload your add-in.</span></span> <span data-ttu-id="fc1d6-106">В этой статье описаны разные способы проверки файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="fc1d6-106">This article describes multiple ways to validate the manifest file and troubleshoot problems with your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="fc1d6-107">Сведения об использовании журнала среды выполнения для устранения неполадок с манифестом надстройки см. в статье [Отладка надстройки с помощью журнала среды выполнения](runtime-logging.md).</span><span class="sxs-lookup"><span data-stu-id="fc1d6-107">For details about using runtime logging to troubleshoot issues with your add-in's manifest, see [Debug your add-in with runtime logging](runtime-logging.md).</span></span>

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a><span data-ttu-id="fc1d6-108">Проверка манифеста с помощью генератора Yeoman для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="fc1d6-108">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="fc1d6-109">Если для создания надстройки использовался [генератор Yeoman для надстроек Office](https://www.npmjs.com/package/generator-office), вы также можете использовать его для проверки файла манифеста проекта.</span><span class="sxs-lookup"><span data-stu-id="fc1d6-109">If you used the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can also use it to validate your project's manifest file.</span></span> <span data-ttu-id="fc1d6-110">Выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="fc1d6-110">Run the following command in the root directory of your project:</span></span>

```command&nbsp;line
npm run validate
```

![GIF-файл с анимацией запуска средства проверки Yo Office в командной строке и получения результатов, которые показывают, что проверка пройдена](../images/yo-office-validator.gif)

> [!NOTE]
> <span data-ttu-id="fc1d6-112">Для доступа к этой функции проект надстройки должен быть создан с помощью [генератора Yeoman](https://www.npmjs.com/package/generator-office) 1.1.17 или более поздней версии.</span><span class="sxs-lookup"><span data-stu-id="fc1d6-112">To have access to this functionality, your add-in project must have been created by using [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) version 1.1.17 or later.</span></span>

## <a name="validate-your-manifest-with-office-addin-manifest"></a><span data-ttu-id="fc1d6-113">Проверка манифеста с помощью office-addin-manifest</span><span class="sxs-lookup"><span data-stu-id="fc1d6-113">Validate your manifest with office-addin-manifest</span></span>

<span data-ttu-id="fc1d6-114">Если для создания надстройки использовался не [генератор Yeoman для надстроек Office](https://www.npmjs.com/package/generator-office), вы можете проверить манифест, используя [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span><span class="sxs-lookup"><span data-stu-id="fc1d6-114">If you didn't use the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can validate the manifest by using [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span></span>

1. <span data-ttu-id="fc1d6-115">Установите [Node.js](https://nodejs.org/download/).</span><span class="sxs-lookup"><span data-stu-id="fc1d6-115">Install [Node.js](https://nodejs.org/download/).</span></span>

2. <span data-ttu-id="fc1d6-116">Выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="fc1d6-116">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="fc1d6-117">Замените `MANIFEST_FILE` на имя файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="fc1d6-117">Replace `MANIFEST_FILE` with the name of the manifest file.</span></span>

    ```command&nbsp;line
    npx office-addin-manifest validate MANIFEST_FILE
    ```

    > [!NOTE]
    > <span data-ttu-id="fc1d6-118">Если эта команда приводит к появлению сообщения об ошибке "Недопустимый синтаксис команды"</span><span class="sxs-lookup"><span data-stu-id="fc1d6-118">If running this command results in the error message "The command syntax is not valid."</span></span> <span data-ttu-id="fc1d6-119">(так как команда `validate` не распознается), выполните следующую команду для проверки манифеста (заменив `MANIFEST_FILE` именем файла манифеста):</span><span class="sxs-lookup"><span data-stu-id="fc1d6-119">(because the `validate` command is not recognized), run the following command to validate the manifest (replacing `MANIFEST_FILE` with the name of the manifest file):</span></span> 
    > 
    > `npx --ignore-existing office-addin-manifest validate MANIFEST_FILE`

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="fc1d6-120">Проверка манифеста на соответствие схеме XML</span><span class="sxs-lookup"><span data-stu-id="fc1d6-120">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="fc1d6-121">Вы можете проверить файл манифеста на соответствие файлам [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas).</span><span class="sxs-lookup"><span data-stu-id="fc1d6-121">You can validate the manifest file against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) files.</span></span> <span data-ttu-id="fc1d6-122">Так вы сможете убедиться в том, что файл манифеста соответствует правильной схеме, включая любые пространства имен для используемых элементов.</span><span class="sxs-lookup"><span data-stu-id="fc1d6-122">This will ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using.</span></span> <span data-ttu-id="fc1d6-123">Если вы скопировали элементы из других примеров манифеста, еще раз проверьте, **включены ли соответствующие пространства имен**.</span><span class="sxs-lookup"><span data-stu-id="fc1d6-123">If you copied elements from other sample manifests double check that you also **include the appropriate namespaces**.</span></span> <span data-ttu-id="fc1d6-124">Для этой проверки можно использовать средство проверки на соответствие схеме XML.</span><span class="sxs-lookup"><span data-stu-id="fc1d6-124">You can use an XML schema validation tool to perform this validation.</span></span>

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="fc1d6-125">Как проверить манифест на соответствие схеме XML с помощью программы командной строки</span><span class="sxs-lookup"><span data-stu-id="fc1d6-125">To use a command-line XML schema validation tool to validate your manifest</span></span>

1. <span data-ttu-id="fc1d6-126">Установите [tar](https://www.gnu.org/software/tar/) и [libxml](http://xmlsoft.org/FAQ.html), если вы еще этого не сделали.</span><span class="sxs-lookup"><span data-stu-id="fc1d6-126">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2. <span data-ttu-id="fc1d6-p106">Выполните указанную ниже команду. Вместо `XSD_FILE` укажите путь к XSD-файлу манифеста, а вместо `XML_FILE` — путь к XML-файлу манифеста.</span><span class="sxs-lookup"><span data-stu-id="fc1d6-p106">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="see-also"></a><span data-ttu-id="fc1d6-129">См. также</span><span class="sxs-lookup"><span data-stu-id="fc1d6-129">See also</span></span>

- [<span data-ttu-id="fc1d6-130">XML-манифест надстройки Office</span><span class="sxs-lookup"><span data-stu-id="fc1d6-130">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="fc1d6-131">Очистка кэша Office</span><span class="sxs-lookup"><span data-stu-id="fc1d6-131">Clear the Office cache</span></span>](clear-cache.md)
- [<span data-ttu-id="fc1d6-132">Отладка надстройки с помощью журнала среды выполнения</span><span class="sxs-lookup"><span data-stu-id="fc1d6-132">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="fc1d6-133">Загрузка неопубликованных надстроек Office для тестирования</span><span class="sxs-lookup"><span data-stu-id="fc1d6-133">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="fc1d6-134">Отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="fc1d6-134">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)