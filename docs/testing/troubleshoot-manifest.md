---
title: Проверка манифеста надстройки Office
description: Узнайте, как проверить манифест надстройки Office с помощью XML-схемы и других средств.
ms.date: 04/16/2020
localization_priority: Normal
ms.openlocfilehash: a6c33603f278e092d5b2b1c85879e03ee10031e3
ms.sourcegitcommit: 803587b324fc8038721709d7db5664025cf03c6b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/17/2020
ms.locfileid: "43547235"
---
# <a name="validate-an-office-add-ins-manifest"></a><span data-ttu-id="60fd7-103">Проверка манифеста надстройки Office</span><span class="sxs-lookup"><span data-stu-id="60fd7-103">Validate an Office Add-in's manifest</span></span>

<span data-ttu-id="60fd7-104">Может потребоваться проверить файл манифеста надстройки, чтобы убедиться в его правильности и полноте.</span><span class="sxs-lookup"><span data-stu-id="60fd7-104">You may want to validate your add-in's manifest file to ensure that it's correct and complete.</span></span> <span data-ttu-id="60fd7-105">Проверка может также выявлять проблемы, которые приводят к появлению ошибки "Манифест надстройки недействителен" при попытке загрузить неопубликованную надстройку.</span><span class="sxs-lookup"><span data-stu-id="60fd7-105">Validation can also identify issues that are causing the error "Your add-in manifest is not valid" when you attempt to sideload your add-in.</span></span> <span data-ttu-id="60fd7-106">В этой статье описаны разные способы проверки файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="60fd7-106">This article describes multiple ways to validate the manifest file.</span></span>

> [!NOTE]
> <span data-ttu-id="60fd7-107">Сведения об использовании журнала среды выполнения для устранения неполадок с манифестом надстройки см. в статье [Отладка надстройки с помощью журнала среды выполнения](runtime-logging.md).</span><span class="sxs-lookup"><span data-stu-id="60fd7-107">For details about using runtime logging to troubleshoot issues with your add-in's manifest, see [Debug your add-in with runtime logging](runtime-logging.md).</span></span>

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a><span data-ttu-id="60fd7-108">Проверка манифеста с помощью генератора Yeoman для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="60fd7-108">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="60fd7-109">Если для создания надстройки использовался [генератор Yeoman для надстроек Office](https://www.npmjs.com/package/generator-office), вы также можете использовать его для проверки файла манифеста проекта.</span><span class="sxs-lookup"><span data-stu-id="60fd7-109">If you used the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can also use it to validate your project's manifest file.</span></span> <span data-ttu-id="60fd7-110">Выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="60fd7-110">Run the following command in the root directory of your project:</span></span>

```command&nbsp;line
npm run validate
```

![GIF-файл с анимацией запуска средства проверки Yo Office в командной строке и получения результатов, которые показывают, что проверка пройдена](../images/yo-office-validator.gif)

> [!NOTE]
> <span data-ttu-id="60fd7-112">Для доступа к этой функции проект надстройки должен быть создан с помощью [генератора Yeoman](https://www.npmjs.com/package/generator-office) 1.1.17 или более поздней версии.</span><span class="sxs-lookup"><span data-stu-id="60fd7-112">To have access to this functionality, your add-in project must have been created by using [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) version 1.1.17 or later.</span></span>

## <a name="validate-your-manifest-with-office-addin-manifest"></a><span data-ttu-id="60fd7-113">Проверка манифеста с помощью office-addin-manifest</span><span class="sxs-lookup"><span data-stu-id="60fd7-113">Validate your manifest with office-addin-manifest</span></span>

<span data-ttu-id="60fd7-114">Если для создания надстройки использовался не [генератор Yeoman для надстроек Office](https://www.npmjs.com/package/generator-office), вы можете проверить манифест, используя [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span><span class="sxs-lookup"><span data-stu-id="60fd7-114">If you didn't use the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can validate the manifest by using [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span></span>

1. <span data-ttu-id="60fd7-115">Установите [Node.js](https://nodejs.org/download/).</span><span class="sxs-lookup"><span data-stu-id="60fd7-115">Install [Node.js](https://nodejs.org/download/).</span></span>

2. <span data-ttu-id="60fd7-116">Выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="60fd7-116">Run the following command in the root directory of your project.</span></span> 

    ```command&nbsp;line
    npm run validate
    ```

    > [!NOTE]
    > <span data-ttu-id="60fd7-117">Если эта команда недоступна или не работает, выполните следующую команду, чтобы принудительно использовать последнюю версию средства Office-ADDIN-MANIFEST (замените `MANIFEST_FILE` на имя файла манифеста):</span><span class="sxs-lookup"><span data-stu-id="60fd7-117">If this command is not available or not working, run the following command instead to force the use of the latest version of the office-addin-manifest tool (replacing `MANIFEST_FILE` with the name of the manifest file):</span></span>
    >
    > ```command&nbsp;line
    > npx --ignore-existing office-addin-manifest validate MANIFEST_FILE
    > ```

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="60fd7-118">Проверка манифеста на соответствие схеме XML</span><span class="sxs-lookup"><span data-stu-id="60fd7-118">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="60fd7-119">Вы можете проверить файл манифеста на соответствие файлам [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).</span><span class="sxs-lookup"><span data-stu-id="60fd7-119">You can validate the manifest file against the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) files.</span></span> <span data-ttu-id="60fd7-120">Так вы сможете убедиться в том, что файл манифеста соответствует правильной схеме, включая любые пространства имен для используемых элементов.</span><span class="sxs-lookup"><span data-stu-id="60fd7-120">This will ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using.</span></span> <span data-ttu-id="60fd7-121">Если вы скопировали элементы из других примеров манифеста, еще раз проверьте, **включены ли соответствующие пространства имен**.</span><span class="sxs-lookup"><span data-stu-id="60fd7-121">If you copied elements from other sample manifests double check that you also **include the appropriate namespaces**.</span></span> <span data-ttu-id="60fd7-122">Для этой проверки можно использовать средство проверки на соответствие схеме XML.</span><span class="sxs-lookup"><span data-stu-id="60fd7-122">You can use an XML schema validation tool to perform this validation.</span></span>

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="60fd7-123">Как проверить манифест на соответствие схеме XML с помощью программы командной строки</span><span class="sxs-lookup"><span data-stu-id="60fd7-123">To use a command-line XML schema validation tool to validate your manifest</span></span>

1. <span data-ttu-id="60fd7-124">Установите [tar](https://www.gnu.org/software/tar/) и [libxml](http://xmlsoft.org/FAQ.html), если вы еще этого не сделали.</span><span class="sxs-lookup"><span data-stu-id="60fd7-124">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2. <span data-ttu-id="60fd7-p104">Выполните указанную ниже команду. Вместо `XSD_FILE` укажите путь к XSD-файлу манифеста, а вместо `XML_FILE` — путь к XML-файлу манифеста.</span><span class="sxs-lookup"><span data-stu-id="60fd7-p104">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="see-also"></a><span data-ttu-id="60fd7-127">См. также</span><span class="sxs-lookup"><span data-stu-id="60fd7-127">See also</span></span>

- [<span data-ttu-id="60fd7-128">XML-манифест надстройки Office</span><span class="sxs-lookup"><span data-stu-id="60fd7-128">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="60fd7-129">Очистка кэша Office</span><span class="sxs-lookup"><span data-stu-id="60fd7-129">Clear the Office cache</span></span>](clear-cache.md)
- [<span data-ttu-id="60fd7-130">Отладка надстройки с помощью журнала среды выполнения</span><span class="sxs-lookup"><span data-stu-id="60fd7-130">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="60fd7-131">Загрузка неопубликованных надстроек Office для тестирования</span><span class="sxs-lookup"><span data-stu-id="60fd7-131">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="60fd7-132">Отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="60fd7-132">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
