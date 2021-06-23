---
title: Отладка надстройки с помощью журнала среды выполнения
description: Узнайте, как использовать журнал среды выполнения для отладки надстройки.
ms.date: 09/23/2020
localization_priority: Normal
ms.openlocfilehash: 3e9a78e6a2f82eca612712f54ac8a700e6d02701
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076415"
---
# <a name="debug-your-add-in-with-runtime-logging"></a><span data-ttu-id="f80a2-103">Отладка надстройки с помощью журнала среды выполнения</span><span class="sxs-lookup"><span data-stu-id="f80a2-103">Debug your add-in with runtime logging</span></span>

<span data-ttu-id="f80a2-104">Вы можете использовать ведение журнала в среде выполнения для отладки манифеста надстройки, а также некоторых ошибок установки.</span><span class="sxs-lookup"><span data-stu-id="f80a2-104">You can use runtime logging to debug your add-in's manifest as well as several installation errors.</span></span> <span data-ttu-id="f80a2-105">Эта функция может помочь вам определять и устранять проблемы с манифестом, которые не обнаруживаются при проверке схемы XSD, например несоответствие идентификаторов ресурсов.</span><span class="sxs-lookup"><span data-stu-id="f80a2-105">This feature can help you identify and fix issues with your manifest that are not detected by XSD schema validation, such as a mismatch between resource IDs.</span></span> <span data-ttu-id="f80a2-106">Ведение журнала в среде выполнения особенно полезно для отладки надстроек, которые добавляют команды и пользовательские функции Excel.</span><span class="sxs-lookup"><span data-stu-id="f80a2-106">Runtime logging is particularly  useful for debugging add-ins that implement add-in commands and Excel custom functions.</span></span>

> [!NOTE]
> <span data-ttu-id="f80a2-107">Функция ведения журнала в настоящее время доступна для Office 2016 или более поздней версии на рабочем столе.</span><span class="sxs-lookup"><span data-stu-id="f80a2-107">The runtime logging feature is currently available for Office 2016 or later on desktop.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f80a2-p102">Ведение журнала в среде выполнения снижает производительность. Включайте его, только когда нужно исправить ошибки в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="f80a2-p102">Runtime Logging affects performance. Turn it on only when you need to debug issues with your add-in manifest.</span></span>

## <a name="use-runtime-logging-from-the-command-line"></a><span data-ttu-id="f80a2-110">Использование журнала в среде выполнения с помощью командной строки</span><span class="sxs-lookup"><span data-stu-id="f80a2-110">Use runtime logging from the command line</span></span>

<span data-ttu-id="f80a2-111">Самый быстрый способ приступить к использованию этого средства ведения журнала — включить ведение журнала в среде выполнения с помощью командной строки.</span><span class="sxs-lookup"><span data-stu-id="f80a2-111">Enabling runtime logging from the command line is the fastest way to use this logging tool.</span></span> <span data-ttu-id="f80a2-112">При этом используется npx (обычно поставляется как часть npm версии 5.2.0 и новее).</span><span class="sxs-lookup"><span data-stu-id="f80a2-112">These use npx, which is provided by default as part of npm@5.2.0+.</span></span> <span data-ttu-id="f80a2-113">Если у вас более ранняя версия [npm](https://www.npmjs.com/), попробуйте воспользоваться инструкциями [Ведение журнала в среде выполнения Windows](#runtime-logging-on-windows) или [Ведение журнала в среде выполнения на компьютере Mac](#runtime-logging-on-mac) либо [установите npx](https://www.npmjs.com/package/npx).</span><span class="sxs-lookup"><span data-stu-id="f80a2-113">If you have an earlier version of [npm](https://www.npmjs.com/), try [Runtime logging on Windows](#runtime-logging-on-windows) or [Runtime logging on Mac](#runtime-logging-on-mac) instructions, or [install npx](https://www.npmjs.com/package/npx).</span></span>

- <span data-ttu-id="f80a2-114">Включение ведения журнала в среде выполнения:</span><span class="sxs-lookup"><span data-stu-id="f80a2-114">To enable runtime logging:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable
    ```

- <span data-ttu-id="f80a2-115">Чтобы включить ведение журнала в среде выполнения только для определенного файла, используйте ту же команду с именем файла:</span><span class="sxs-lookup"><span data-stu-id="f80a2-115">To enable runtime logging only for a specific file, use the same command with a filename:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --enable [filename.txt]
    ```

- <span data-ttu-id="f80a2-116">Отключение ведения журнала в среде выполнения:</span><span class="sxs-lookup"><span data-stu-id="f80a2-116">To disable runtime logging:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --disable
    ```

- <span data-ttu-id="f80a2-117">Определение, включено ли ведение журнала в среде выполнения:</span><span class="sxs-lookup"><span data-stu-id="f80a2-117">To display whether runtime logging is enabled:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log
    ```

- <span data-ttu-id="f80a2-118">Отображение справки в командной строке для ведения журнала в среде выполнения:</span><span class="sxs-lookup"><span data-stu-id="f80a2-118">To display help within the command line for runtime logging:</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings runtime-log --help
    ```

## <a name="runtime-logging-on-windows"></a><span data-ttu-id="f80a2-119">Ведение журнала в среде выполнения в Windows</span><span class="sxs-lookup"><span data-stu-id="f80a2-119">Runtime logging on Windows</span></span>

1. <span data-ttu-id="f80a2-120">Убедитесь, что у вас установлена сборка Office 2016 **16.0.7019** или выше.</span><span class="sxs-lookup"><span data-stu-id="f80a2-120">Make sure that you are running Office 2016 desktop build **16.0.7019** or later.</span></span>

2. <span data-ttu-id="f80a2-121">Добавьте раздел реестра `RuntimeLogging` в раздел `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span><span class="sxs-lookup"><span data-stu-id="f80a2-121">Add the `RuntimeLogging` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`.</span></span>

    [!include[Developer registry key](../includes/developer-registry-key.md)]


3. <span data-ttu-id="f80a2-122">В качестве значения параметра **RuntimeLogging** по умолчанию задайте полный путь к файлу, в который будет записываться журнал.</span><span class="sxs-lookup"><span data-stu-id="f80a2-122">Set the default value of the **RuntimeLogging** key to the full path of the file where you want the log to be written.</span></span> <span data-ttu-id="f80a2-123">Пример приведен в архиве [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span><span class="sxs-lookup"><span data-stu-id="f80a2-123">For an example, see [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip).</span></span>

    > [!NOTE]
    > <span data-ttu-id="f80a2-124">Необходим готовый каталог, в котором будет создан файл журнала, и соответствующее разрешение на запись.</span><span class="sxs-lookup"><span data-stu-id="f80a2-124">The directory in which the log file will be written must already exist, and you must have write permissions to it.</span></span>

<span data-ttu-id="f80a2-p105">Ниже показано, как должен выглядеть реестр. Чтобы отключить функцию, удалите из реестра раздел `RuntimeLogging`.</span><span class="sxs-lookup"><span data-stu-id="f80a2-p105">The following image shows what the registry should look like. To turn the feature off, remove the `RuntimeLogging` key from the registry.</span></span>

![Снимок экрана редактора реестра с ключом реестра RuntimeLogging.](../images/runtime-logging-registry.png)

## <a name="runtime-logging-on-mac"></a><span data-ttu-id="f80a2-128">Ведение журнала в среде выполнения на компьютере Mac</span><span class="sxs-lookup"><span data-stu-id="f80a2-128">Runtime logging on Mac</span></span>

1. <span data-ttu-id="f80a2-129">Убедитесь, что у вас установлена классическая сборка Office 2016 **16.27** (19071500) или более поздней версии.</span><span class="sxs-lookup"><span data-stu-id="f80a2-129">Make sure that you are running Office 2016 desktop build **16.27** (19071500) or later.</span></span>

2. <span data-ttu-id="f80a2-130">Откройте приложение **Терминал** и настройте параметры ведения журнала в среде выполнения с помощью команды `defaults`:</span><span class="sxs-lookup"><span data-stu-id="f80a2-130">Open **Terminal** and set a runtime logging preference by using the `defaults` command:</span></span>

    ```command&nbsp;line
    defaults write <bundle id> CEFRuntimeLoggingFile -string <file_name>
    ```

    <span data-ttu-id="f80a2-131">`<bundle id>` указывает, для какого узла требуется включить ведение журнала в среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="f80a2-131">`<bundle id>` identifies which the host for which to enable runtime logging.</span></span> <span data-ttu-id="f80a2-132">`<file_name>` — это имя текстового файла, в который будет записан журнал.</span><span class="sxs-lookup"><span data-stu-id="f80a2-132">`<file_name>` is the name of the text file to which the log will be written.</span></span>

    <span data-ttu-id="f80a2-133">Установите одно из следующих значений, чтобы включить журнал времени работы `<bundle id>` для соответствующего приложения:</span><span class="sxs-lookup"><span data-stu-id="f80a2-133">Set `<bundle id>` to one of the following values to enable runtime logging for the corresponding application:</span></span>

    - `com.microsoft.Word`
    - `com.microsoft.Excel`
    - `com.microsoft.Powerpoint`
    - `com.microsoft.Outlook`

<span data-ttu-id="f80a2-134">В следующем примере включается ведение журнала в среде выполнения в Word, а затем открывается файл журнала:</span><span class="sxs-lookup"><span data-stu-id="f80a2-134">The following example enables runtime logging for Word and then opens the log file:</span></span>

```command&nbsp;line
defaults write com.microsoft.Word CEFRuntimeLoggingFile -string "runtime_logs.txt"
open ~/library/Containers/com.microsoft.Word/Data/runtime_logs.txt
```

> [!NOTE]
> <span data-ttu-id="f80a2-135">Чтобы включить ведение журнала в среде выполнения, потребуется перезапустить Office после выполнения команды `defaults`.</span><span class="sxs-lookup"><span data-stu-id="f80a2-135">You'll need to restart Office after running the `defaults` command to enable runtime logging.</span></span>

<span data-ttu-id="f80a2-136">Чтобы отключить ведение журнала в среде выполнения, используйте команду `defaults delete`:</span><span class="sxs-lookup"><span data-stu-id="f80a2-136">To turn off runtime logging, use the `defaults delete` command:</span></span>

```command&nbsp;line
defaults delete <bundle id> CEFRuntimeLoggingFile
```

<span data-ttu-id="f80a2-137">В следующем примере отключается ведение журнала в среде выполнения для Word.</span><span class="sxs-lookup"><span data-stu-id="f80a2-137">The following example will turn off runtime logging for Word:</span></span>

```command&nbsp;line
defaults delete com.microsoft.Word CEFRuntimeLoggingFile
```

## <a name="use-runtime-logging-to-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="f80a2-138">Используйте журнал среды выполнения для устранения неполадок манифеста</span><span class="sxs-lookup"><span data-stu-id="f80a2-138">Use runtime logging to troubleshoot issues with your manifest</span></span>

<span data-ttu-id="f80a2-139">Чтобы устранить проблемы с загрузкой надстройки, используя журнал среды выполнения:</span><span class="sxs-lookup"><span data-stu-id="f80a2-139">To use runtime logging to troubleshoot issues loading an add-in:</span></span>

1. <span data-ttu-id="f80a2-140">[Загрузите неопубликованную надстройку](sideload-office-add-ins-for-testing.md) для тестирования.</span><span class="sxs-lookup"><span data-stu-id="f80a2-140">[Sideload your add-in](sideload-office-add-ins-for-testing.md) for testing.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f80a2-141">Рекомендуем загружать только тестируемую надстройку, чтобы уменьшить количество сообщений в файле журнала.</span><span class="sxs-lookup"><span data-stu-id="f80a2-141">We recommend that you sideload only the add-in that you are testing to minimize the number of messages in the log file.</span></span>

2. <span data-ttu-id="f80a2-142">Если ничего не происходит и надстройка не отображается в диалоговом окне надстроек, откройте файл журнала.</span><span class="sxs-lookup"><span data-stu-id="f80a2-142">If nothing happens and you don't see your add-in (and it's not appearing in the add-ins dialog box), open the log file.</span></span>

3. <span data-ttu-id="f80a2-p107">Выполните в этом файле поиск по идентификатору надстройки, определенному в манифесте. В файле журнала этот идентификатор отмечен как `SolutionId`.</span><span class="sxs-lookup"><span data-stu-id="f80a2-p107">Search the log file for your add-in ID, which you define in your manifest. In the log file, this ID is labeled `SolutionId`.</span></span>

## <a name="known-issues-with-runtime-logging"></a><span data-ttu-id="f80a2-145">Известные проблемы с ведением журнала в среде выполнения</span><span class="sxs-lookup"><span data-stu-id="f80a2-145">Known issues with runtime logging</span></span>

<span data-ttu-id="f80a2-p108">В файле журнала могут встречаться непонятные или неправильно классифицированные сообщения. Например:</span><span class="sxs-lookup"><span data-stu-id="f80a2-p108">You might see messages in the log file that are confusing or that are classified incorrectly. For example:</span></span>

- <span data-ttu-id="f80a2-148">сообщение `Medium Current host not in add-in's host list` с дополнением `Unexpected Parsed manifest targeting different host` неправильно классифицируется как ошибка.</span><span class="sxs-lookup"><span data-stu-id="f80a2-148">The message `Medium Current host not in add-in's host list` followed by `Unexpected Parsed manifest targeting different host` is incorrectly classified as an error.</span></span>

- <span data-ttu-id="f80a2-149">Если появится сообщение `Unexpected Add-in is missing required manifest fields    DisplayName`, не содержащее SolutionId, то ошибка, скорее всего, не связана с надстройкой, отладка которой выполняется.</span><span class="sxs-lookup"><span data-stu-id="f80a2-149">If you see the message `Unexpected Add-in is missing required manifest fields    DisplayName` and it doesn't contain a SolutionId, the error is most likely not related to the add-in you are debugging.</span></span>

- <span data-ttu-id="f80a2-p109">Все сообщения `Monitorable` являются ожидаемыми ошибками с точки зрения системы. Иногда они указывают на проблему с манифестом, например опечатку в элементе, которая была пропущена, но не привела к сбою.</span><span class="sxs-lookup"><span data-stu-id="f80a2-p109">Any `Monitorable` messages are expected errors from a system point of view. Sometimes they indicate an issue with your manifest, such as a misspelled element that was skipped but didn't cause the manifest to fail.</span></span>

## <a name="see-also"></a><span data-ttu-id="f80a2-152">См. также</span><span class="sxs-lookup"><span data-stu-id="f80a2-152">See also</span></span>

- [<span data-ttu-id="f80a2-153">XML-манифест надстройки Office</span><span class="sxs-lookup"><span data-stu-id="f80a2-153">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="f80a2-154">Проверка манифеста надстройки Office</span><span class="sxs-lookup"><span data-stu-id="f80a2-154">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="f80a2-155">Очистка кэша Office</span><span class="sxs-lookup"><span data-stu-id="f80a2-155">Clear the Office cache</span></span>](clear-cache.md)
- [<span data-ttu-id="f80a2-156">Загрузка неопубликованных надстроек Office для тестирования</span><span class="sxs-lookup"><span data-stu-id="f80a2-156">Sideload Office Add-ins for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="f80a2-157">Отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="f80a2-157">Debug Office Add-ins</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
