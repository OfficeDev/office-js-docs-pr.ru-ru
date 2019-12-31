---
title: Отладка надстройки с помощью журнала среды выполнения
description: Узнайте, как использовать журнал среды выполнения для отладки надстройки.
ms.date: 12/31/2019
localization_priority: Priority
ms.openlocfilehash: d69811963caf2b6d48b400ac9744d38e53859167
ms.sourcegitcommit: d5ac9284d1e96dc91a9168d7641e44d88535e1a7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/31/2019
ms.locfileid: "40915085"
---
# <a name="debug-your-add-in-with-runtime-logging"></a><span data-ttu-id="0490c-103">Отладка надстройки с помощью журнала среды выполнения</span><span class="sxs-lookup"><span data-stu-id="0490c-103">Debug your add-in with runtime logging</span></span>

<span data-ttu-id="0490c-104">Вы можете использовать ведение журнала в среде выполнения для отладки манифеста надстройки, а также некоторых ошибок установки.</span><span class="sxs-lookup"><span data-stu-id="0490c-104">You can use runtime logging to debug your add-in's manifest as well as several installation errors.</span></span> <span data-ttu-id="0490c-105">Эта функция может помочь вам определять и устранять проблемы с манифестом, которые не обнаруживаются при проверке схемы XSD, например несоответствие идентификаторов ресурсов.</span><span class="sxs-lookup"><span data-stu-id="0490c-105">This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs.</span></span> <span data-ttu-id="0490c-106">Ведение журнала в среде выполнения особенно полезно для отладки надстроек, которые добавляют команды и пользовательские функции Excel.</span><span class="sxs-lookup"><span data-stu-id="0490c-106">Runtime logging is particularly  useful for debugging add-ins that implement add-in commands and Excel custom functions.</span></span>   

> [!NOTE]
> <span data-ttu-id="0490c-107">В настоящее время функция ведения журнала в среде выполнения доступна для классических приложений Office 2016.</span><span class="sxs-lookup"><span data-stu-id="0490c-107">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0490c-108">Ведение журнала в среде выполнения сказывается на производительности.</span><span class="sxs-lookup"><span data-stu-id="0490c-108">Runtime Logging affects performance.</span></span> <span data-ttu-id="0490c-109">Включайте его, только если требуется устранить неполадки, связанные с манифестом надстройки.</span><span class="sxs-lookup"><span data-stu-id="0490c-109">Turn it on only when you need to debug issues with your add-in manifest.</span></span>

## <a name="use-runtime-logging-from-the-command-line"></a><span data-ttu-id="0490c-110">Использование журнала в среде выполнения с помощью командной строки</span><span class="sxs-lookup"><span data-stu-id="0490c-110">Use runtime logging from the command line</span></span>

<span data-ttu-id="0490c-111">Самый быстрый способ приступить к использованию этого средства ведения журнала — включить ведение журнала в среде выполнения с помощью командной строки.</span><span class="sxs-lookup"><span data-stu-id="0490c-111">Enabling runtime logging from the command line is the fastest way to use this logging tool.</span></span> <span data-ttu-id="0490c-112">При этом используется npx (обычно поставляется как часть npm версии 5.2.0 и новее).</span><span class="sxs-lookup"><span data-stu-id="0490c-112">These use npx, which is provided by default as part of npm@5.2.0+.</span></span> <span data-ttu-id="0490c-113">Если у вас более ранняя версия [npm](https://www.npmjs.com/), попробуйте воспользоваться инструкциями [Ведение журнала в среде выполнения Windows](#runtime-logging-on-windows) или [Ведение журнала в среде выполнения на компьютере Mac](#runtime-logging-on-mac) либо [установите npx](https://www.npmjs.com/package/npx).</span><span class="sxs-lookup"><span data-stu-id="0490c-113">If you have an earlier version of [npm](https://www.npmjs.com/), try [Runtime logging on Windows](#runtime-logging-on-windows) or [Runtime logging on Mac](#runtime-logging-on-mac) instructions, or [install npx](https://www.npmjs.com/package/npx).</span></span>

- <span data-ttu-id="0490c-114">Включение ведения журнала в среде выполнения:</span><span class="sxs-lookup"><span data-stu-id="0490c-114">To enable runtime logging:</span></span>
    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable
    ```
- <span data-ttu-id="0490c-115">Чтобы включить ведение журнала в среде выполнения только для определенного файла, используйте ту же команду с именем файла:</span><span class="sxs-lookup"><span data-stu-id="0490c-115">To enable runtime logging only for a specific file, use the same command with a filename:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable [filename.txt]
    ```

- <span data-ttu-id="0490c-116">Отключение ведения журнала в среде выполнения:</span><span class="sxs-lookup"><span data-stu-id="0490c-116">To disable runtime logging:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --disable
    ```

- <span data-ttu-id="0490c-117">Определение, включено ли ведение журнала в среде выполнения:</span><span class="sxs-lookup"><span data-stu-id="0490c-117">To display whether runtime logging is enabled:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log
    ```

- <span data-ttu-id="0490c-118">Отображение справки в командной строке для ведения журнала в среде выполнения:</span><span class="sxs-lookup"><span data-stu-id="0490c-118">To display help within the command line for runtime logging:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --help
    ```

## <a name="runtime-logging-on-windows"></a><span data-ttu-id="0490c-119">Ведение журнала в среде выполнения в Windows</span><span class="sxs-lookup"><span data-stu-id="0490c-119">Runtime logging on Windows</span></span>

1. <span data-ttu-id="0490c-120">Убедитесь, что у вас установлена сборка Office 2016 **16.0.7019** или выше.</span><span class="sxs-lookup"><span data-stu-id="0490c-120">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span> 

2. <span data-ttu-id="0490c-121">Добавьте раздел реестра `RuntimeLogging` в раздел `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span><span class="sxs-lookup"><span data-stu-id="0490c-121">Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="0490c-122">Если ключа (папки) `Developer` еще нет в разделе `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, создайте его, выполнив следующие действия:</span><span class="sxs-lookup"><span data-stu-id="0490c-122">If the `Developer` key (folder) does not already exist under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, complete the following steps to create it:</span></span> 
    > 1. <span data-ttu-id="0490c-123">Щелкните правой кнопкой мыши ключ (папку) **WEF** и выберите **Создать** > **Ключ**.</span><span class="sxs-lookup"><span data-stu-id="0490c-123">Right-click the **WEF** key (folder) and select **New** > **Key**.</span></span>
    > 2. <span data-ttu-id="0490c-124">Назовите новый ключ **Разработчик**.</span><span class="sxs-lookup"><span data-stu-id="0490c-124">Name the new key **Developer**.</span></span>

3. <span data-ttu-id="0490c-125">В качестве значения параметра **RuntimeLogging** по умолчанию задайте полный путь к файлу, в который будет записываться журнал.</span><span class="sxs-lookup"><span data-stu-id="0490c-125">Set the default value of the **RuntimeLogging** key to the full path of the file where you want the log to be written.</span></span> <span data-ttu-id="0490c-126">Пример приведен в архиве [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span><span class="sxs-lookup"><span data-stu-id="0490c-126">For an example, see [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="0490c-127">Необходим готовый каталог, в котором будет создан файл журнала, и соответствующее разрешение на запись.</span><span class="sxs-lookup"><span data-stu-id="0490c-127">The directory in which the log file will be written must already exist, and you must have write permissions to it.</span></span> 
 
<span data-ttu-id="0490c-p105">Ниже показано, как должен выглядеть реестр. Чтобы отключить функцию, удалите из реестра раздел `RuntimeLogging`.</span><span class="sxs-lookup"><span data-stu-id="0490c-p105">The following image shows what the registry should look like. To turn the feature off, remove the `RuntimeLogging` key from the registry.</span></span> 

![Снимок экрана: редактор реестра с разделом RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)

## <a name="runtime-logging-on-mac"></a><span data-ttu-id="0490c-131">Ведение журнала в среде выполнения на компьютере Mac</span><span class="sxs-lookup"><span data-stu-id="0490c-131">Runtime logging on Mac</span></span>

1. <span data-ttu-id="0490c-132">Убедитесь, что у вас установлена классическая сборка Office 2016 **16.27** (19071500) или более поздней версии.</span><span class="sxs-lookup"><span data-stu-id="0490c-132">Make sure that you are running Office 2016 desktop build **16.27** (19071500) or later.</span></span>

2. <span data-ttu-id="0490c-133">Откройте приложение **Терминал** и настройте параметры ведения журнала в среде выполнения с помощью команды `defaults`:</span><span class="sxs-lookup"><span data-stu-id="0490c-133">Open **Terminal** and set a runtime logging preference by using the `defaults` command:</span></span>
    
    ```command&nbsp;line
    defaults write <bundle id> CEFRuntimeLoggingFile -string <file_name>
    ```

    <span data-ttu-id="0490c-134">`<bundle id>` указывает, для какого узла требуется включить ведение журнала в среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="0490c-134">`<bundle id>` identifies which the host for which to enable runtime logging.</span></span> <span data-ttu-id="0490c-135">`<file_name>` — это имя текстового файла, в который будет записан журнал.</span><span class="sxs-lookup"><span data-stu-id="0490c-135">`<file_name>` is the name of the text file to which the log will be written.</span></span>

    <span data-ttu-id="0490c-136">Чтобы включить ведение журнала в среде выполнения для соответствующего узла, присвойте параметру `<bundle id>` одно из следующих значений:</span><span class="sxs-lookup"><span data-stu-id="0490c-136">Set `<bundle id>` to one of the following values to enable runtime logging for the corresponding host:</span></span>

    - `com.microsoft.Word`
    - `com.microsoft.Excel`
    - `com.microsoft.Powerpoint`
    - `com.microsoft.Outlook`

<span data-ttu-id="0490c-137">В следующем примере включается ведение журнала в среде выполнения в Word, а затем открывается файл журнала:</span><span class="sxs-lookup"><span data-stu-id="0490c-137">The following example enables runtime logging for Word and then opens the log file:</span></span>

```command&nbsp;line
defaults write com.microsoft.Word CEFRuntimeLoggingFile -string "runtime_logs.txt"
open ~/library/Containers/com.microsoft.Word/Data/runtime_logs.txt
```

> [!NOTE] 
> <span data-ttu-id="0490c-138">Чтобы включить ведение журнала в среде выполнения, потребуется перезапустить Office после выполнения команды `defaults`.</span><span class="sxs-lookup"><span data-stu-id="0490c-138">You'll need to restart Office after running the `defaults` command to enable runtime logging.</span></span>

<span data-ttu-id="0490c-139">Чтобы отключить ведение журнала в среде выполнения, используйте команду `defaults delete`:</span><span class="sxs-lookup"><span data-stu-id="0490c-139">To turn off runtime logging, use the `defaults delete` command:</span></span>

```command&nbsp;line
defaults delete <bundle id> CEFRuntimeLoggingFile
```

<span data-ttu-id="0490c-140">В следующем примере отключается ведение журнала в среде выполнения для Word.</span><span class="sxs-lookup"><span data-stu-id="0490c-140">The following example will turn off runtime logging for Word:</span></span>

```command&nbsp;line
defaults delete com.microsoft.Word CEFRuntimeLoggingFile
```

## <a name="use-runtime-logging-to-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="0490c-141">Используйте журнал среды выполнения для устранения неполадок манифеста</span><span class="sxs-lookup"><span data-stu-id="0490c-141">Use runtime logging to troubleshoot issues with your manifest</span></span>

<span data-ttu-id="0490c-142">Чтобы устранить проблемы с загрузкой надстройки, используя журнал среды выполнения:</span><span class="sxs-lookup"><span data-stu-id="0490c-142">To use runtime logging to troubleshoot issues loading an add-in:</span></span>
 
1. <span data-ttu-id="0490c-143">[Загрузите неопубликованную надстройку](sideload-office-add-ins-for-testing.md) для тестирования.</span><span class="sxs-lookup"><span data-stu-id="0490c-143">[Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="0490c-144">Рекомендуем загружать только тестируемую надстройку, чтобы уменьшить количество сообщений в файле журнала.</span><span class="sxs-lookup"><span data-stu-id="0490c-144">We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.</span></span>

2. <span data-ttu-id="0490c-145">Если ничего не происходит и надстройка не отображается в диалоговом окне надстроек, откройте файл журнала.</span><span class="sxs-lookup"><span data-stu-id="0490c-145">If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.</span></span>

3. <span data-ttu-id="0490c-p107">Выполните в этом файле поиск по идентификатору надстройки, определенному в манифесте. В файле журнала этот идентификатор отмечен как `SolutionId`.</span><span class="sxs-lookup"><span data-stu-id="0490c-p107">Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.</span></span> 

<span data-ttu-id="0490c-p108">В приведенном ниже примере файл журнала определяет элемент управления, указывающий на несуществующий файл ресурсов. В этом примере необходимо исправить опечатку в манифесте или добавить недостающий ресурс.</span><span class="sxs-lookup"><span data-stu-id="0490c-p108">In the following example, the log file identifies a control that points to a resource file that doesn't exist. For this example, the fix would be to correct the typo in the manifest or to add the missing resource.</span></span>

![Снимок экрана с файлом журнала, содержащим запись, которая указывает на несуществующий идентификатор ресурса.](http://i.imgur.com/f8bouLA.png) 

## <a name="known-issues-with-runtime-logging"></a><span data-ttu-id="0490c-151">Известные проблемы с ведением журнала в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="0490c-151">Known issues with runtime logging</span></span>

<span data-ttu-id="0490c-p109">В файле журнала могут встречаться непонятные или неправильно классифицированные сообщения. Например:</span><span class="sxs-lookup"><span data-stu-id="0490c-p109">You might see messages in the log file that are confusing or that are classified incorrectly. For example:</span></span>

- <span data-ttu-id="0490c-154">сообщение `Medium Current host not in add-in's host list` с дополнением `Unexpected Parsed manifest targeting different host` неправильно классифицируется как ошибка.</span><span class="sxs-lookup"><span data-stu-id="0490c-154">The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.</span></span>

- <span data-ttu-id="0490c-155">Если появится сообщение `Unexpected Add-in is missing required manifest fields DisplayName`, не содержащее SolutionId, то ошибка, скорее всего, не связана с надстройкой, отладка которой выполняется.</span><span class="sxs-lookup"><span data-stu-id="0490c-155">If you see the message `Unexpected Add-in is missing required manifest fields DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.</span></span> 

- <span data-ttu-id="0490c-p110">Все сообщения `Monitorable` являются ожидаемыми ошибками с точки зрения системы. Иногда они указывают на проблему с манифестом, например опечатку в элементе, которая была пропущена, но не привела к сбою.</span><span class="sxs-lookup"><span data-stu-id="0490c-p110">Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.</span></span> 

## <a name="see-also"></a><span data-ttu-id="0490c-158">См. также</span><span class="sxs-lookup"><span data-stu-id="0490c-158">See also</span></span>

- [<span data-ttu-id="0490c-159">XML-манифест надстройки Office</span><span class="sxs-lookup"><span data-stu-id="0490c-159">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="0490c-160">Проверка манифеста надстройки Office</span><span class="sxs-lookup"><span data-stu-id="0490c-160">Validate an Office Add-in manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="0490c-161">Очистка кэша Office</span><span class="sxs-lookup"><span data-stu-id="0490c-161">Clear the Office cache</span></span>](clear-cache.md)
- [<span data-ttu-id="0490c-162">Загрузка неопубликованных надстроек Office для тестирования</span><span class="sxs-lookup"><span data-stu-id="0490c-162">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="0490c-163">Отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="0490c-163">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)