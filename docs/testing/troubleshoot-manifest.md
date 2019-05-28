---
title: Проверка манифеста и устранение связанных с ним неполадок
description: Используйте эти методы для проверки манифеста надстройки Office.
ms.date: 05/21/2019
localization_priority: Priority
ms.openlocfilehash: 5b9bd22ad724bac68587a41ad56f4290f3a6edbd
ms.sourcegitcommit: adaee1329ae9bb69e49bde7f54a4c0444c9ba642
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/24/2019
ms.locfileid: "34432266"
---
# <a name="validate-and-troubleshoot-issues-with-your-manifest"></a>Проверка манифеста и устранение связанных с ним неполадок

Проверить манифест надстройки Office и устранить связанные с ним неполадки можно с помощью указанных ниже методов. 

- [Проверка манифеста с помощью средства проверки надстроек Office](#validate-your-manifest-with-the-office-add-in-validator)   
- [Проверка манифеста на соответствие схеме XML](#validate-your-manifest-against-the-xml-schema)
- [Проверка манифеста с помощью генератора Yeoman для надстроек Office](#validate-your-manifest-with-the-yeoman-generator-for-office-add-ins)
- [Отладка надстройки с помощью журнала среды выполнения](#use-runtime-logging-to-debug-your-add-in)


## <a name="validate-your-manifest-with-the-office-add-in-validator"></a>Проверка манифеста с помощью средства проверки надстроек Office

Чтобы убедиться, что файл манифеста правильно и полностью описывает надстройку Office, проверьте его с помощью [средства проверки надстроек Office](https://github.com/OfficeDev/office-addin-validator).

### <a name="to-use-the-office-add-in-validator-to-validate-your-manifest"></a>Как проверить манифест с помощью средства проверки надстроек Office

1. Установите [Node.js](https://nodejs.org/download/). 

2. Откройте командную строку или терминал от имени администратора и глобально установите средство проверки надстроек, используя следующую команду:

    ```command&nbsp;line
    npm install -g office-addin-validator
    ```
    
    > [!NOTE]
    > Если у вас уже установлено приложение Yo Office, обновите его до последней версии, при этом средство проверки будет установлено в виде зависимости.

3. Выполните приведенную ниже команду для проверки манифеста. Вместо файла MANIFEST.XML укажите путь к XML-файлу манифеста.

    ```command&nbsp;line
    validate-office-addin MANIFEST.XML
    ```

## <a name="validate-your-manifest-against-the-xml-schema"></a>Проверка манифеста на соответствие схеме XML

Проверьте файл манифеста на соответствие правильной схеме, в том числе пространства имен для используемых элементов. Если вы скопировали элементы из других примеров манифеста, еще раз проверьте, **включены ли соответствующие пространства имен**. Вы можете проверить манифест, используя файлы [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas). Для этой проверки можно использовать средство проверки на соответствие схеме XML. 



### <a name="to-use-a-command-line-xml-schema-validation-tool-to-validate-your-manifest"></a>Как проверить манифест на соответствие схеме XML с помощью программы командной строки

1.  Установите [tar](https://www.gnu.org/software/tar/) и [libxml](http://xmlsoft.org/FAQ.html), если вы еще этого не сделали.

2.  Выполните указанную ниже команду. Вместо `XSD_FILE` укажите путь к XSD-файлу манифеста, а вместо `XML_FILE` — путь к XML-файлу манифеста.
    
    ```command&nbsp;line
    xmllint --noout --schema XSD_FILE XML_FILE
    ```

## <a name="validate-your-manifest-with-the-yeoman-generator-for-office-add-ins"></a>Проверка манифеста с помощью генератора Yeoman для надстроек Office

Если вы создали надстройку Office, используя [генератора Yeoman](https://www.npmjs.com/package/generator-office), убедитесь, что файл манифеста соответствует правильной схеме, выполнив следующую команду в корневом каталоге проекта:

```command&nbsp;line
npm run validate
```

![GIF-файл с анимацией запуска средства проверки Yo Office в командной строке и получения результатов, которые показывают, что проверка пройдена](../images/yo-office-validator.gif)

> [!NOTE]
> Для доступа к этой функции проект надстройки должен быть создан с помощью [генератора Yeoman](https://www.npmjs.com/package/generator-office) 1.1.17 или более поздней версии.

## <a name="use-runtime-logging-to-debug-your-add-in"></a>Отладка надстройки с помощью журнала среды выполнения 

Вы можете использовать ведение журнала в среде выполнения для отладки манифеста надстройки, а также некоторых ошибок установки. Эта функция может помочь вам определять и устранять проблемы с манифестом, которые не обнаруживаются при проверке схемы XSD, например несоответствие идентификаторов ресурсов. Ведение журнала в среде выполнения особенно полезно для отладки надстроек, которые добавляют команды и пользовательские функции Excel.   

> [!NOTE]
> В настоящее время функция ведения журнала в среде выполнения доступна для классических приложений Office 2016.

### <a name="to-turn-on-runtime-logging"></a>Как включить ведение журнала в среде выполнения

> [!IMPORTANT]
> Ведение журнала в среде выполнения снижает производительность. Включайте его, только когда нужно исправить ошибки в манифесте надстройки.

Чтобы включить ведение журнала в среде выполнения:

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
Удалите содержимое папки `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`. 

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a>Для iOS
Для принудительной перезагрузки вызовите метод JavaScript `window.location.reload(true)` в надстройке. Вы также можете переустановить Office.

## <a name="see-also"></a>См. также

- [XML-манифест надстройки Office](../develop/add-in-manifests.md)
- [Загрузка неопубликованных надстроек Office для тестирования](sideload-office-add-ins-for-testing.md)
- [Отладка надстроек Office](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
