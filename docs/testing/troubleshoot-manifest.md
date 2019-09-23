---
title: Проверка манифеста и устранение связанных с ним неполадок
description: Используйте эти методы для проверки манифеста надстройки Office.
ms.date: 09/18/2019
localization_priority: Priority
ms.openlocfilehash: c320c05b944bba9e24a4d3c0e5ef514ac13cc3c6
ms.sourcegitcommit: a0257feabcfe665061c14b8bdb70cf82f7aca414
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/18/2019
ms.locfileid: "37035338"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a>Проверка манифеста и устранение связанных с ним неполадок

Может потребоваться проверить файл манифеста надстройки, чтобы убедиться в его правильности и полноте. Проверка может также выявлять проблемы, которые приводят к появлению ошибки "Манифест надстройки недействителен" при попытке загрузить неопубликованную надстройку. В этой статье описано несколько способов проверки файла манифеста и устранения связанных с надстройкой неполадок.

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a>Проверка манифеста с помощью генератора Yeoman для надстроек Office

Если для создания надстройки использовался [генератор Yeoman для надстроек Office](https://www.npmjs.com/package/generator-office), вы также можете использовать его для проверки файла манифеста проекта. Выполните следующую команду в корневом каталоге своего проекта.

```command&nbsp;line
npm run validate
```

![GIF-файл с анимацией запуска средства проверки Yo Office в командной строке и получения результатов, которые показывают, что проверка пройдена](../images/yo-office-validator.gif)

> [!NOTE]
> Для доступа к этой функции проект надстройки должен быть создан с помощью [генератора Yeoman](https://www.npmjs.com/package/generator-office) 1.1.17 или более поздней версии.

## <a name="validate-your-manifest-with-office-addin-manifest"></a>Проверка манифеста с помощью office-addin-manifest

Если для создания надстройки использовался не [генератор Yeoman для надстроек Office](https://www.npmjs.com/package/generator-office), вы можете проверить манифест, используя [office-addin-manifest](https://www.npmjs.com/package/office-addin-manifest).

1. Установите [Node.js](https://nodejs.org/download/).

2. Выполните следующую команду в корневом каталоге своего проекта. Замените `MANIFEST_FILE` на имя файла манифеста.

    ```command&nbsp;line
    npx office-addin-manifest validate MANIFEST_FILE
    ```

    > [!NOTE]
    > Если эта команда приводит к появлению сообщения об ошибке "Недопустимый синтаксис команды" (так как команда `validate` не распознается), выполните следующую команду для проверки манифеста (заменив `MANIFEST_FILE` именем файла манифеста): 
    > 
    > `npx --ignore-existing office-addin-manifest validate MANIFEST_FILE`

## <a name="validate-your-manifest-against-the-xml-schema"></a>Проверка манифеста на соответствие схеме XML

Вы можете проверить файл манифеста на соответствие файлам [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas). Так вы сможете убедиться в том, что файл манифеста соответствует правильной схеме, включая любые пространства имен для используемых элементов. Если вы скопировали элементы из других примеров манифеста, еще раз проверьте, **включены ли соответствующие пространства имен**. Для этой проверки можно использовать средство проверки на соответствие схеме XML.

### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a>Как проверить манифест на соответствие схеме XML с помощью программы командной строки

1. Установите [tar](https://www.gnu.org/software/tar/) и [libxml](http://xmlsoft.org/FAQ.html), если вы еще этого не сделали.

2. Выполните указанную ниже команду. Вместо `XSD_FILE` укажите путь к XSD-файлу манифеста, а вместо `XML_FILE` — путь к XML-файлу манифеста.
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="use-runtime-logging-to-debug-your-add-in"></a>Отладка надстройки с помощью журнала среды выполнения

Вы можете использовать ведение журнала в среде выполнения для отладки манифеста надстройки, а также некоторых ошибок установки. Эта функция может помочь вам определять и устранять проблемы с манифестом, которые не обнаруживаются при проверке схемы XSD, например несоответствие идентификаторов ресурсов. Ведение журнала в среде выполнения особенно полезно для отладки надстроек, которые добавляют команды и пользовательские функции Excel.   

> [!NOTE]
> В настоящее время функция ведения журнала в среде выполнения доступна для классических приложений Office 2016.

> [!IMPORTANT]
> Ведение журнала в среде выполнения сказывается на производительности. Включайте его, только если требуется устранить неполадки, связанные с манифестом надстройки.

### <a name="runtime-logging-on-windows"></a>Ведение журнала в среде выполнения в Windows

1. Убедитесь, что у вас установлена сборка Office 2016 **16.0.7019** или выше. 

2. Добавьте раздел реестра `RuntimeLogging` в раздел `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\`. 

    > [!NOTE]
    > Если ключа (папки) `Developer` еще нет в разделе `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\`, создайте его, выполнив следующие действия: 
    > 1. Щелкните правой кнопкой мыши ключ (папку) **WEF** и выберите **Создать** > **Ключ**.
    > 2. Назовите новый ключ **Разработчик**.

3. В качестве значения по умолчанию задайте полный путь к файлу, в который будет записываться журнал. Пример приведен в архиве [EnableRuntimeLogging.zip](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/raw/master/Tools/RuntimeLogging/EnableRuntimeLogging.zip). 

    > [!NOTE]
    > Необходим готовый каталог, в котором будет создан файл журнала, и соответствующее разрешение на запись. 
 
Ниже показано, как должен выглядеть реестр. Чтобы отключить функцию, удалите из реестра раздел `RuntimeLogging`. 

![Снимок экрана: редактор реестра с разделом RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)

### <a name="runtime-logging-on-mac"></a>Ведение журнала в среде выполнения на компьютере Mac

1. Убедитесь, что у вас установлена классическая сборка Office 2016 **16.27** (19071500) или более поздней версии.

2. Откройте приложение **Терминал** и настройте параметры ведения журнала в среде выполнения с помощью команды `defaults`:
    
    ```command&nbsp;line
    defaults write <bundle id> CEFRuntimeLoggingFile -string <file_name>
    ```

    `<bundle id>` указывает, для какого узла требуется включить ведение журнала в среде выполнения. `<file_name>` — это имя текстового файла, в который будет записан журнал.

    Чтобы включить ведение журнала в среде выполнения для соответствующего узла, присвойте параметру `<bundle id>` одно из следующих значений:

    - `com.microsoft.Word`
    - `com.microsoft.Excel`
    - `com.microsoft.Powerpoint`
    - `com.microsoft.Outlook`

В следующем примере включается ведение журнала в среде выполнения в Word, а затем открывается файл журнала:

```command&nbsp;line
defaults write com.microsoft.Word CEFRuntimeLoggingFile -string "runtime_logs.txt"
open ~/library/Containers/com.microsoft.Word/Data/runtime_logs.txt
```

> [!NOTE] 
> Чтобы включить ведение журнала в среде выполнения, потребуется перезапустить Office после выполнения команды `defaults`.

Чтобы отключить ведение журнала в среде выполнения, используйте команду `defaults delete`:

```command&nbsp;line
defaults delete <bundle id> CEFRuntimeLoggingFile
```

В следующем примере отключается ведение журнала в среде выполнения для Word.

```command&nbsp;line
defaults delete com.microsoft.Word CEFRuntimeLoggingFile
```

### <a name="to-troubleshoot-issues-with-your-manifest"></a>Как устранить проблемы с манифестом

Чтобы устранить проблемы с загрузкой надстройки, используя журнал среды выполнения:
 
1. [Загрузите неопубликованную надстройку](sideload-office-add-ins-for-testing.md) для тестирования. 

    > [!NOTE]
    > Рекомендуем загружать только тестируемую надстройку, чтобы уменьшить количество сообщений в файле журнала.

2. Если ничего не происходит и надстройка не отображается в диалоговом окне надстроек, откройте файл журнала.

3. Выполните в этом файле поиск по идентификатору надстройки, определенному в манифесте. В файле журнала этот идентификатор отмечен как `SolutionId`. 

В приведенном ниже примере файл журнала определяет элемент управления, указывающий на несуществующий файл ресурсов. В этом примере необходимо исправить опечатку в манифесте или добавить недостающий ресурс.

![Снимок экрана с файлом журнала, содержащим запись, которая указывает на несуществующий идентификатор ресурса.](http://i.imgur.com/f8bouLA.png) 

### <a name="known-issues-with-runtime-logging"></a>Известные проблемы с ведением журнала в среде выполнения

В файле журнала могут встречаться непонятные или неправильно классифицированные сообщения. Например:

- сообщение `Medium Current host not in add-in's host list` с дополнением `Unexpected Parsed manifest targeting different host` неправильно классифицируется как ошибка.

- Если появится сообщение `Unexpected Add-in is missing required manifest fields DisplayName`, не содержащее SolutionId, то ошибка, скорее всего, не связана с надстройкой, отладка которой выполняется. 

- Все сообщения `Monitorable` являются ожидаемыми ошибками с точки зрения системы. Иногда они указывают на проблему с манифестом, например опечатку в элементе, которая была пропущена, но не привела к сбою. 

## <a name="clear-the-office-cache"></a>Очистка кэша Office

Если внесенные в манифест изменения (например, имена значков кнопок на ленте или текст команд надстроек) не вступили в силу, попробуйте очистить кэш Office на своем компьютере. 

#### <a name="for-windows"></a>Для Windows
Удалите содержимое папки `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.

#### <a name="for-mac"></a>Для Mac

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a>Для iOS
Для принудительной перезагрузки вызовите метод JavaScript `window.location.reload(true)` в надстройке. Вы также можете переустановить Office.

## <a name="see-also"></a>См. также

- [XML-манифест надстройки Office](../develop/add-in-manifests.md)
- [Загрузка неопубликованных надстроек Office для тестирования](sideload-office-add-ins-for-testing.md)
- [Отладка надстроек Office](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
