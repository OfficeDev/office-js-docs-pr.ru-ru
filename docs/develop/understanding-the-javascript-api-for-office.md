---
title: Общие сведения об интерфейсе API JavaScript для Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: fef2cdad69408f099296461066f1ea380e3b118b
ms.sourcegitcommit: bc68b4cf811b45e8b8d1cbd7c8d2867359ab671b
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2018
ms.locfileid: "21703814"
---
# <a name="understanding-the-javascript-api-for-office"></a>Общие сведения об интерфейсе API JavaScript для Office

В этой статье можно узнать об интерфейсе API JavaScript для Office и о том, как его использовать. Справочные сведения см. в разделе [API JavaScript для Office](https://dev.office.com/reference/add-ins/javascript-api-for-office). О том, как обновить файлы проекта Visual Studio до последней версии API JavaScript для Office, см. в статье [Обновление версии API JavaScript для Office и файлов схемы манифеста](update-your-javascript-api-for-office-and-manifest-schema-version.md).

> [!NOTE]
> Если вы планируете [опубликовать](../publish/publish.md) надстройку в AppSource и сделать ее доступной в интерфейсе Office, убедитесь, что она соответствует [политикам проверки AppSource](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). Например, чтобы пройти проверку, надстройка работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) и на [странице со сведениями о доступности и ведущих приложениях для надстроек Office](../overview/office-add-in-availability.md). 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a>Ссылки на библиотеку API JavaScript для Office в надстройке

Библиотека [API JavaScript для Office](https://dev.office.com/reference/add-ins/javascript-api-for-office) состоит из файла Office.js и связанных JS-файлов ведущего приложения, например Excel-15.js и Outlook-15.js. Простейший способ сослаться на API — использовать нашу сеть доставки содержимого (CDN), добавив следующий код `<script>` в тег `<head>` страницы:  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

Это приведет к скачиванию и кэшированию файлов JavaScript API для Office при первой загрузке надстройки, чтобы убедиться, что она использует актуальную реализацию Office.js и сопутствующих файлов для указанной версии.

Подробные сведения о CDN Office.js, включая способы управления версиями и обратной совместимостью, см. в статье [Указание ссылок на библиотеку API JavaScript для Office из сети доставки содержимого (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="initializing-your-add-in"></a>Инициализация надстройки

**Область применения:** все типы надстроек

Библиотека Office.js включает событие инициализации, которое вызывается, когда API полностью загружено и готово к взаимодействию с пользователем. С помощью обработчика события **initialize** можно реализовать распространенные сценарии инициализации надстройки, например предложение выбрать некоторые ячейки в Excel и вставку диаграммы, инициализированной с помощью выбранных значений. Кроме того, с помощью обработчика события initialize можно инициализировать другую пользовательскую логику для надстройки, например установку привязок, запросы на значения параметров надстройки по умолчанию и т. д.

В простейшем случае событие initialize будет выглядеть как в следующем примере:     

```js
Office.initialize = function () { };
```
При использовании дополнительных платформ JavaScript, включающих собственный обработчик событий инициализации или тесты, их следует размещать внутри события Office.initialize. Например, ссылка на функцию [JQuery](https://jquery.com) `$(document).ready()` должна выглядеть следующим образом:

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {        
        // The document is ready
    });
  };
```

На всех страницах надстроек Office необходимо назначить обработчик события initialize, **Office.initialize**. Если не назначить обработчик события, при запуске надстройки может возникнуть ошибка. Кроме того, если пользователь попробует использовать надстройку с веб-клиентом Office Online, например Excel Online, PowerPoint Online или Outlook Web App, произойдет сбой. Если вам не нужен код инициализации, то функция, назначенная событию **Office.initialize**, может не содержать кода, как показано в первом из приведенных выше примеров.

Дополнительные сведения о последовательности событий при инициализации надстройки см. в статье [Загрузка модели DOM и среды выполнения](loading-the-dom-and-runtime-environment.md).

#### <a name="initialization-reason"></a>Причина инициализации
Для надстроек области задач и контентных надстроек Office.initialize обеспечивает дополнительный параметр _reason_. Этот параметр можно использовать для определения способа, каким надстройка была добавлена в текущий документ. Это поможет обеспечить разную логику в тех случаях, когда надстройка вставляется впервые или когда она уже существует в документе. 

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
        switch (reason) {
            case 'inserted': console.log('The add-in was just inserted.');
            case 'documentOpened': console.log('The add-in is already part of the document.');
        }
    });
 };
```
Дополнительные сведения см. в статьях [Событие Office.initialize Event](https://dev.office.com/reference/add-ins/shared/office.initialize) и [Перечисление InitializationReason](https://dev.office.com/reference/add-ins/shared/initializationreason-enumeration). 

## <a name="office-javascript-api-object-model"></a>Объектная модель API JavaScript для Office

После инициализации надстройка может взаимодействовать с узлом (например, Excel, Outlook). Более подробную информацию о конкретных шаблонах использования см. на странице [Объектная модель API Office JavaScript](/office-javascript-api-object-model.md). Существует также подробная справочная документация как для узлов[общих API-интерфейсов](https://dev.office.com/reference/add-ins/javascript-api-for-office), так и для конкретных узлов.

## <a name="api-support-matrix"></a>Матрица поддержки API


В этой таблице представлены API и функции, поддерживаемые всеми типами надстроек (контентными, области задач и Outlook), и приложения Office, в которых они могут работать, когда вы указываете ведущие приложения Office, поддерживаемые вашей надстройкой с помощью [схемы манифестов надстроек версии 1.1 и функций, поддерживаемых API JavaScript для Office версии 1.1](update-your-javascript-api-for-office-and-manifest-schema-version.md).


|||||||||
|:-----|:-----|:-----:|:-----:|:-----:|:-----:|:-----:|:-----:|
||**Имя узла**|База данных|Книга|Почтовый ящик|Презентация|Документ|Проект|
||**Поддерживаемые** **ведущие приложения**|Веб-приложения Access|Excel,<br/>Excel Online|Outlook,<br/>веб-приложение Outlook,<br/>Outlook Web App для устройств|PowerPoint,<br/>PowerPoint Online|Word|Проект|
|**Поддерживаемые типы надстроек**|Содержимое|Да|Да||Да|||
||Область задач||Да||Да|Да|Да|
||Outlook|||Да||||
|**Поддерживаемые функции API**|Чтение и запись текста||Да||Да|Да|Да<br/>(только для чтения)|
||Матрица чтения и записи||Да|||Да||
||Таблица чтения и записи||Да|||Да||
||Чтение и запись HTML|||||Да||
||Чтение и запись<br/>Office Open XML|||||Да||
||Чтение свойств задач, ресурсов, представлений и полей||||||Да|
||События изменения выделенного фрагмента||Да|||Да||
||Загрузка всего документа||||Да|Да||
||Привязки и их события|Да<br/>(только полный и частичные привязки таблиц)|Да|||Да||
||Чтение и запись настраиваемых частей XML|||||Да||
||Сохранение данных состояния надстройки (параметры)|Да<br/>(на ведущую надстройку)|Да<br/>(на документ)|Да<br/>(на почтовый ящик)|Да<br/>(на документ)|Да<br/>(на документ)||
||События изменения параметров|Да|Да||Да|Да||
||События получения активного режима просмотра<br/>и изменения представления||||Да|||
||Переход к расположениям<br/>в документе||Да||Да|Да||
||Активация в зависимости от контекста<br/>с помощью правил и RegEx|||Да||||
||Чтение свойств элемента|||Да||||
||Чтение профиля пользователя|||Да||||
||Получение вложений|||Да||||
||Получение токена удостоверения пользователя|||Да||||
||Вызов веб-служб Exchange|||Да||||
