---
title: Проверка манифеста и устранение связанных с ним неполадок
description: Используйте эти методы для проверки манифеста надстройки Office.
ms.date: 10/29/2019
localization_priority: Priority
ms.openlocfilehash: c1af6308a975bf9204a519e21f828454d286aa19
ms.sourcegitcommit: 59d29d01bce7543ebebf86e5a86db00cf54ca14a
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/01/2019
ms.locfileid: "37924810"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="feed7-103">Проверка манифеста и устранение связанных с ним неполадок</span><span class="sxs-lookup"><span data-stu-id="feed7-103">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="feed7-104">Может потребоваться проверить файл манифеста надстройки, чтобы убедиться в его правильности и полноте.</span><span class="sxs-lookup"><span data-stu-id="feed7-104">You may want to validate your add-in's manifest file to ensure that it's correct and complete.</span></span> <span data-ttu-id="feed7-105">Проверка может также выявлять проблемы, которые приводят к появлению ошибки "Манифест надстройки недействителен" при попытке загрузить неопубликованную надстройку.</span><span class="sxs-lookup"><span data-stu-id="feed7-105">Validation can also identify issues that are causing the error "Your add-in manifest is not valid" when you attempt to sideload your add-in.</span></span> <span data-ttu-id="feed7-106">В этой статье описано несколько способов проверки файла манифеста и устранения связанных с надстройкой неполадок.</span><span class="sxs-lookup"><span data-stu-id="feed7-106">This article describes multiple ways to validate the manifest file and troubleshoot problems with your add-in.</span></span>

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a><span data-ttu-id="feed7-107">Проверка манифеста с помощью генератора Yeoman для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="feed7-107">Validate your manifest with the Yeoman generator for Office Add-ins</span></span>

<span data-ttu-id="feed7-108">Если для создания надстройки использовался [генератор Yeoman для надстроек Office](https://www.npmjs.com/package/generator-office), вы также можете использовать его для проверки файла манифеста проекта.</span><span class="sxs-lookup"><span data-stu-id="feed7-108">If you used the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can also use it to validate your project's manifest file.</span></span> <span data-ttu-id="feed7-109">Выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="feed7-109">Run the following command in the root directory of your project:</span></span>

```command&nbsp;line
npm run validate
```

![GIF-файл с анимацией запуска средства проверки Yo Office в командной строке и получения результатов, которые показывают, что проверка пройдена](../images/yo-office-validator.gif)

> [!NOTE]
> <span data-ttu-id="feed7-111">Для доступа к этой функции проект надстройки должен быть создан с помощью [генератора Yeoman](https://www.npmjs.com/package/generator-office) 1.1.17 или более поздней версии.</span><span class="sxs-lookup"><span data-stu-id="feed7-111">To have access to this functionality, your add-in project must have been created by using [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) version 1.1.17 or later.</span></span>

## <a name="validate-your-manifest-with-office-addin-manifest"></a><span data-ttu-id="feed7-112">Проверка манифеста с помощью office-addin-manifest</span><span class="sxs-lookup"><span data-stu-id="feed7-112">Validate your manifest with office-addin-manifest</span></span>

<span data-ttu-id="feed7-113">Если для создания надстройки использовался не [генератор Yeoman для надстроек Office](https://www.npmjs.com/package/generator-office), вы можете проверить манифест, используя [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span><span class="sxs-lookup"><span data-stu-id="feed7-113">If you didn't use the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to create your add-in, you can validate the manifest by using [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).</span></span>

1. <span data-ttu-id="feed7-114">Установите [Node.js](https://nodejs.org/download/).</span><span class="sxs-lookup"><span data-stu-id="feed7-114">Install [Node.js](https://nodejs.org/download/).</span></span>

2. <span data-ttu-id="feed7-115">Выполните следующую команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="feed7-115">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="feed7-116">Замените `MANIFEST_FILE` на имя файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="feed7-116">Replace `MANIFEST_FILE` with the name of the manifest file.</span></span>

    ```command&nbsp;line
    npx office-addin-manifest validate MANIFEST_FILE
    ```

    > [!NOTE]
    > <span data-ttu-id="feed7-117">Если эта команда приводит к появлению сообщения об ошибке "Недопустимый синтаксис команды"</span><span class="sxs-lookup"><span data-stu-id="feed7-117">If running this command results in the error message "The command syntax is not valid."</span></span> <span data-ttu-id="feed7-118">(так как команда `validate` не распознается), выполните следующую команду для проверки манифеста (заменив `MANIFEST_FILE` именем файла манифеста):</span><span class="sxs-lookup"><span data-stu-id="feed7-118">(because the `validate` command is not recognized), run the following command to validate the manifest (replacing `MANIFEST_FILE` with the name of the manifest file):</span></span> 
    > 
    > `npx --ignore-existing office-addin-manifest validate MANIFEST_FILE`

## <a name="validate-your-manifest-against-the-xml-schema"></a><span data-ttu-id="feed7-119">Проверка манифеста на соответствие схеме XML</span><span class="sxs-lookup"><span data-stu-id="feed7-119">Validate your manifest against the XML schema</span></span>

<span data-ttu-id="feed7-120">Вы можете проверить файл манифеста на соответствие файлам [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas).</span><span class="sxs-lookup"><span data-stu-id="feed7-120">You can validate the manifest file against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas) files.</span></span> <span data-ttu-id="feed7-121">Так вы сможете убедиться в том, что файл манифеста соответствует правильной схеме, включая любые пространства имен для используемых элементов.</span><span class="sxs-lookup"><span data-stu-id="feed7-121">This will ensure that the manifest file follows the correct schema, including any namespaces for the elements you are using.</span></span> <span data-ttu-id="feed7-122">Если вы скопировали элементы из других примеров манифеста, еще раз проверьте, **включены ли соответствующие пространства имен**.</span><span class="sxs-lookup"><span data-stu-id="feed7-122">If you copied elements from other sample manifests double check that you also **include the appropriate namespaces**.</span></span> <span data-ttu-id="feed7-123">Для этой проверки можно использовать средство проверки на соответствие схеме XML.</span><span class="sxs-lookup"><span data-stu-id="feed7-123">You can use an XML schema validation tool to perform this validation.</span></span>

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a><span data-ttu-id="feed7-124">Как проверить манифест на соответствие схеме XML с помощью программы командной строки</span><span class="sxs-lookup"><span data-stu-id="feed7-124">To use a command-line XML schema validation tool to validate your manifest</span></span>

1. <span data-ttu-id="feed7-125">Установите [tar](https://www.gnu.org/software/tar/) и [libxml](http://xmlsoft.org/FAQ.html), если вы еще этого не сделали.</span><span class="sxs-lookup"><span data-stu-id="feed7-125">Install [tar](https://www.gnu.org/software/tar/) and [libxml](http://xmlsoft.org/FAQ.html), if you haven't already.</span></span>

2. <span data-ttu-id="feed7-p106">Выполните указанную ниже команду. Вместо `XSD_FILE` укажите путь к XSD-файлу манифеста, а вместо `XML_FILE` — путь к XML-файлу манифеста.</span><span class="sxs-lookup"><span data-stu-id="feed7-p106">Run the following command. Replace `XSD_FILE` with the path to the manifest XSD file, and replace `XML_FILE` with the path to the manifest XML file.</span></span>
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="use-runtime-logging-to-debug-your-add-in"></a><span data-ttu-id="feed7-128">Отладка надстройки с помощью журнала среды выполнения</span><span class="sxs-lookup"><span data-stu-id="feed7-128">Use runtime logging to debug your add-in</span></span>

<span data-ttu-id="feed7-129">Вы можете использовать ведение журнала в среде выполнения для отладки манифеста надстройки, а также некоторых ошибок установки.</span><span class="sxs-lookup"><span data-stu-id="feed7-129">You can use runtime logging to debug your add-in's manifest as well as several installation errors.</span></span> <span data-ttu-id="feed7-130">Эта функция может помочь вам определять и устранять проблемы с манифестом, которые не обнаруживаются при проверке схемы XSD, например несоответствие идентификаторов ресурсов.</span><span class="sxs-lookup"><span data-stu-id="feed7-130">This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs.</span></span> <span data-ttu-id="feed7-131">Ведение журнала в среде выполнения особенно полезно для отладки надстроек, которые добавляют команды и пользовательские функции Excel.</span><span class="sxs-lookup"><span data-stu-id="feed7-131">Runtime logging is particularly  useful for debugging add-ins that implement add-in commands and Excel custom functions.</span></span>   

> [!NOTE]
> <span data-ttu-id="feed7-132">В настоящее время функция ведения журнала в среде выполнения доступна для классических приложений Office 2016.</span><span class="sxs-lookup"><span data-stu-id="feed7-132">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="feed7-133">Ведение журнала в среде выполнения сказывается на производительности.</span><span class="sxs-lookup"><span data-stu-id="feed7-133">Runtime Logging affects performance.</span></span> <span data-ttu-id="feed7-134">Включайте его, только если требуется устранить неполадки, связанные с манифестом надстройки.</span><span class="sxs-lookup"><span data-stu-id="feed7-134">Turn it on only when you need to debug issues with your add-in manifest.</span></span>

### <a name="use-runtime-logging-from-the-command-line"></a><span data-ttu-id="feed7-135">Использование журнала в среде выполнения с помощью командной строки</span><span class="sxs-lookup"><span data-stu-id="feed7-135">Use runtime logging from the command line</span></span>

<span data-ttu-id="feed7-136">Самый быстрый способ приступить к использованию этого средства ведения журнала — включить ведение журнала в среде выполнения с помощью командной строки.</span><span class="sxs-lookup"><span data-stu-id="feed7-136">Enabling runtime logging from the command line is the fastest way to use this logging tool.</span></span> <span data-ttu-id="feed7-137">При этом используется npx (обычно поставляется как часть npm версии 5.2.0 и новее).</span><span class="sxs-lookup"><span data-stu-id="feed7-137">These use npx, which is provided by default as part of npm@5.2.0+.</span></span> <span data-ttu-id="feed7-138">Если у вас более ранняя версия [npm](https://www.npmjs.com/), попробуйте воспользоваться инструкциями [Ведение журнала в среде выполнения Windows](#runtime-logging-on-windows) или [Ведение журнала в среде выполнения на компьютере Mac](#runtime-logging-on-mac) либо [установите npx](https://www.npmjs.com/package/npx).</span><span class="sxs-lookup"><span data-stu-id="feed7-138">If you have an earlier version of [npm](https://www.npmjs.com/), try [Runtime logging on Windows](#runtime-logging-on-windows) or [Runtime logging on Mac](#runtime-logging-on-mac) instructions, or [install npx](https://www.npmjs.com/package/npx).</span></span>

- <span data-ttu-id="feed7-139">Включение ведения журнала в среде выполнения:</span><span class="sxs-lookup"><span data-stu-id="feed7-139">To enable AD FS logging</span></span>
    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable
    ```
- <span data-ttu-id="feed7-140">Чтобы включить ведение журнала в среде выполнения только для определенного файла, используйте ту же команду с именем файла:</span><span class="sxs-lookup"><span data-stu-id="feed7-140">To enable runtime logging only for a specific file, use the same command with a filename:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable [filename.txt]
    ```

- <span data-ttu-id="feed7-141">Отключение ведения журнала в среде выполнения:</span><span class="sxs-lookup"><span data-stu-id="feed7-141">To disable runtime logging:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --disable
    ```

- <span data-ttu-id="feed7-142">Определение, включено ли ведение журнала в среде выполнения:</span><span class="sxs-lookup"><span data-stu-id="feed7-142">To display whether runtime logging is enabled:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log
    ```

- <span data-ttu-id="feed7-143">Отображение справки в командной строке для ведения журнала в среде выполнения:</span><span class="sxs-lookup"><span data-stu-id="feed7-143">To display help within the command line for runtime logging:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --help
    ```

### <a name="runtime-logging-on-windows"></a><span data-ttu-id="feed7-144">Ведение журнала в среде выполнения в Windows</span><span class="sxs-lookup"><span data-stu-id="feed7-144">Runtime logging on Windows</span></span>

1. <span data-ttu-id="feed7-145">Убедитесь, что у вас установлена сборка Office 2016 **16.0.7019** или выше.</span><span class="sxs-lookup"><span data-stu-id="feed7-145">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span> 

2. <span data-ttu-id="feed7-146">Добавьте раздел реестра `RuntimeLogging` в раздел `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span><span class="sxs-lookup"><span data-stu-id="feed7-146">Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="feed7-147">Если ключа (папки) `Developer` еще нет в разделе `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, создайте его, выполнив следующие действия:</span><span class="sxs-lookup"><span data-stu-id="feed7-147">If the `Developer` key (folder) does not already exist under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, complete the following steps to create it:</span></span> 
    > 1. <span data-ttu-id="feed7-148">Щелкните правой кнопкой мыши ключ (папку) **WEF** и выберите **Создать** > **Ключ**.</span><span class="sxs-lookup"><span data-stu-id="feed7-148">Right-click the **WEF** key (folder) and select **New** > **Key**.</span></span>
    > 2. <span data-ttu-id="feed7-149">Назовите новый ключ **Разработчик**.</span><span class="sxs-lookup"><span data-stu-id="feed7-149">Name the new key **Developer**.</span></span>

3. <span data-ttu-id="feed7-150">В качестве значения параметра **RuntimeLogging** по умолчанию задайте полный путь к файлу, в который будет записываться журнал.</span><span class="sxs-lookup"><span data-stu-id="feed7-150">Set the default value of the key to the full path of the file where you want the log to be written.</span></span> <span data-ttu-id="feed7-151">Пример приведен в архиве [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span><span class="sxs-lookup"><span data-stu-id="feed7-151">For an example, see [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="feed7-152">Необходим готовый каталог, в котором будет создан файл журнала, и соответствующее разрешение на запись.</span><span class="sxs-lookup"><span data-stu-id="feed7-152">The directory in which the log file will be written must already exist, and you must have write permissions to it.</span></span> 
 
<span data-ttu-id="feed7-p111">Ниже показано, как должен выглядеть реестр. Чтобы отключить функцию, удалите из реестра раздел `RuntimeLogging`.</span><span class="sxs-lookup"><span data-stu-id="feed7-p111">The following image shows what the registry should look like. To turn the feature off, remove the `RuntimeLogging` key from the registry.</span></span> 

![Снимок экрана: редактор реестра с разделом RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)

### <a name="runtime-logging-on-mac"></a><span data-ttu-id="feed7-156">Ведение журнала в среде выполнения на компьютере Mac</span><span class="sxs-lookup"><span data-stu-id="feed7-156">Runtime logging on Mac</span></span>

1. <span data-ttu-id="feed7-157">Убедитесь, что у вас установлена классическая сборка Office 2016 **16.27** (19071500) или более поздней версии.</span><span class="sxs-lookup"><span data-stu-id="feed7-157">Make sure that you are running Office 2016 desktop build **16.27** (19071500) or later.</span></span>

2. <span data-ttu-id="feed7-158">Откройте приложение **Терминал** и настройте параметры ведения журнала в среде выполнения с помощью команды `defaults`:</span><span class="sxs-lookup"><span data-stu-id="feed7-158">Open **Terminal** and set a runtime logging preference by using the `defaults` command:</span></span>
    
    ```command&nbsp;line
    defaults write <bundle id> CEFRuntimeLoggingFile -string <file_name>
    ```

    <span data-ttu-id="feed7-159">`<bundle id>` указывает, для какого узла требуется включить ведение журнала в среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="feed7-159">`<bundle id>` identifies which the host for which to enable runtime logging.</span></span> <span data-ttu-id="feed7-160">`<file_name>` — это имя текстового файла, в который будет записан журнал.</span><span class="sxs-lookup"><span data-stu-id="feed7-160">`<file_name>` is the name of the text file to which the log will be written.</span></span>

    <span data-ttu-id="feed7-161">Чтобы включить ведение журнала в среде выполнения для соответствующего узла, присвойте параметру `<bundle id>` одно из следующих значений:</span><span class="sxs-lookup"><span data-stu-id="feed7-161">Set `<bundle id>` to one of the following values to enable runtime logging for the corresponding host:</span></span>

    - `com.microsoft.Word`
    - `com.microsoft.Excel`
    - `com.microsoft.Powerpoint`
    - `com.microsoft.Outlook`

<span data-ttu-id="feed7-162">В следующем примере включается ведение журнала в среде выполнения в Word, а затем открывается файл журнала:</span><span class="sxs-lookup"><span data-stu-id="feed7-162">The following example enables runtime logging for Word and then opens the log file:</span></span>

```command&nbsp;line
defaults write com.microsoft.Word CEFRuntimeLoggingFile -string "runtime_logs.txt"
open ~/library/Containers/com.microsoft.Word/Data/runtime_logs.txt
```

> [!NOTE] 
> <span data-ttu-id="feed7-163">Чтобы включить ведение журнала в среде выполнения, потребуется перезапустить Office после выполнения команды `defaults`.</span><span class="sxs-lookup"><span data-stu-id="feed7-163">You'll need to restart Office after running the `defaults` command to enable runtime logging.</span></span>

<span data-ttu-id="feed7-164">Чтобы отключить ведение журнала в среде выполнения, используйте команду `defaults delete`:</span><span class="sxs-lookup"><span data-stu-id="feed7-164">To turn off runtime logging, use the `defaults delete` command:</span></span>

```command&nbsp;line
defaults delete <bundle id> CEFRuntimeLoggingFile
```

<span data-ttu-id="feed7-165">В следующем примере отключается ведение журнала в среде выполнения для Word.</span><span class="sxs-lookup"><span data-stu-id="feed7-165">The following example will turn off runtime logging for Word:</span></span>

```command&nbsp;line
defaults delete com.microsoft.Word CEFRuntimeLoggingFile
```

### <a name="to-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="feed7-166">Как устранить проблемы с манифестом</span><span class="sxs-lookup"><span data-stu-id="feed7-166">To troubleshoot issues with your manifest</span></span>

<span data-ttu-id="feed7-167">Чтобы устранить проблемы с загрузкой надстройки, используя журнал среды выполнения:</span><span class="sxs-lookup"><span data-stu-id="feed7-167">To use runtime logging to troubleshoot issues loading an add-in:</span></span>
 
1. <span data-ttu-id="feed7-168">[Загрузите неопубликованную надстройку](sideload-office-add-ins-for-testing.md) для тестирования.</span><span class="sxs-lookup"><span data-stu-id="feed7-168">[Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="feed7-169">Рекомендуем загружать только тестируемую надстройку, чтобы уменьшить количество сообщений в файле журнала.</span><span class="sxs-lookup"><span data-stu-id="feed7-169">We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.</span></span>

2. <span data-ttu-id="feed7-170">Если ничего не происходит и надстройка не отображается в диалоговом окне надстроек, откройте файл журнала.</span><span class="sxs-lookup"><span data-stu-id="feed7-170">If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.</span></span>

3. <span data-ttu-id="feed7-p113">Выполните в этом файле поиск по идентификатору надстройки, определенному в манифесте. В файле журнала этот идентификатор отмечен как `SolutionId`.</span><span class="sxs-lookup"><span data-stu-id="feed7-p113">Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.</span></span> 

<span data-ttu-id="feed7-p114">В приведенном ниже примере файл журнала определяет элемент управления, указывающий на несуществующий файл ресурсов. В этом примере необходимо исправить опечатку в манифесте или добавить недостающий ресурс.</span><span class="sxs-lookup"><span data-stu-id="feed7-p114">In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.</span></span>

![Снимок экрана с файлом журнала, содержащим запись, которая указывает на несуществующий идентификатор ресурса.](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a><span data-ttu-id="feed7-176">Известные проблемы с ведением журнала в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="feed7-176">Known issues with runtime logging</span></span>

<span data-ttu-id="feed7-p115">В файле журнала могут встречаться непонятные или неправильно классифицированные сообщения. Например:</span><span class="sxs-lookup"><span data-stu-id="feed7-p115">You might see messages in the log file that are confusing or that are classified incorrectly. For example:</span></span>

- <span data-ttu-id="feed7-179">сообщение `Medium Current host not in add-in's host list` с дополнением `Unexpected Parsed manifest targeting different host` неправильно классифицируется как ошибка.</span><span class="sxs-lookup"><span data-stu-id="feed7-179">The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.</span></span>

- <span data-ttu-id="feed7-180">Если появится сообщение `Unexpected Add-in is missing required manifest fields DisplayName`, не содержащее SolutionId, то ошибка, скорее всего, не связана с надстройкой, отладка которой выполняется.</span><span class="sxs-lookup"><span data-stu-id="feed7-180">If you see the message `Unexpected Add-in is missing required manifest fields DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.</span></span> 

- <span data-ttu-id="feed7-p116">Все сообщения `Monitorable` являются ожидаемыми ошибками с точки зрения системы. Иногда они указывают на проблему с манифестом, например опечатку в элементе, которая была пропущена, но не привела к сбою.</span><span class="sxs-lookup"><span data-stu-id="feed7-p116">Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.</span></span> 

## <a name="clear-the-office-cache"></a><span data-ttu-id="feed7-183">Очистка кэша Office</span><span class="sxs-lookup"><span data-stu-id="feed7-183">Clear the Office cache</span></span>

<span data-ttu-id="feed7-184">Если внесенные в манифест изменения (например, имена значков кнопок на ленте или текст команд надстроек) не вступили в силу, попробуйте очистить кэш Office на своем компьютере.</span><span class="sxs-lookup"><span data-stu-id="feed7-184">If changes you've made in the manifest, such as file names of ribbon button icons or text of add-in commands, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="feed7-185">Для Windows</span><span class="sxs-lookup"><span data-stu-id="feed7-185">For Windows:</span></span>
<span data-ttu-id="feed7-186">Удалите содержимое папки `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="feed7-186">Delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="feed7-187">Для Mac</span><span class="sxs-lookup"><span data-stu-id="feed7-187">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="feed7-188">Для iOS</span><span class="sxs-lookup"><span data-stu-id="feed7-188">For iOS:</span></span>
<span data-ttu-id="feed7-p117">Для принудительной перезагрузки вызовите метод JavaScript `window.location.reload(true)` в надстройке. Вы также можете переустановить Office.</span><span class="sxs-lookup"><span data-stu-id="feed7-p117">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="feed7-191">См. также</span><span class="sxs-lookup"><span data-stu-id="feed7-191">See also</span></span>

- [<span data-ttu-id="feed7-192">XML-манифест надстройки Office</span><span class="sxs-lookup"><span data-stu-id="feed7-192">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="feed7-193">Загрузка неопубликованных надстроек Office для тестирования</span><span class="sxs-lookup"><span data-stu-id="feed7-193">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="feed7-194">Отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="feed7-194">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
