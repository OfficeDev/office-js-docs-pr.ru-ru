---
title: Создание ссылок на библиотеку API JavaScript для Office
description: Узнайте, как ссылаться на библиотеку API JavaScript для Office и определение типов в надстройке.
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 9f7753b24e0a5861778b09ea93fecdc26fd2ca96
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325159"
---
# <a name="referencing-the-office-javascript-api-library"></a>Создание ссылок на библиотеку API JavaScript для Office

Библиотека [API JavaScript для Office](../reference/javascript-api-for-office.md) предоставляет API, которые надстройка может использовать для взаимодействия с ведущим приложением Office. Самый простой способ добавить ссылку на библиотеку — использовать сеть доставки содержимого (CDN), добавив следующий `<script>` тег в `<head>` раздел страницы HTML:  

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

Это приведет к скачиванию и кэшированию файлов API JavaScript для Office при первом запуске надстройки, чтобы убедиться в том, что используется самая последняя реализация Office. js и связанных с ней файлов для указанной версии.

> [!IMPORTANT]
> Необходимо ссылаться на API JavaScript для Office из `<head>` раздела страницы, чтобы убедиться, что API полностью инициализирован до элементов основного текста. Ведущим приложениям Office необходимо, чтобы надстройки инициализировались в течение 5 секунд после активации. Если надстройка не активируется в этом пороговом значении, она будет объявлена без ответа, а пользователю будет выведено сообщение об ошибке.

## <a name="api-versioning-and-backward-compatibility"></a>Управление версиями и обратная совместимость API

В предыдущем фрагменте кода HTML ( `/1/` перед `office.js` в URL-адресе CDN) указывает последний добавочный выпуск в версии 1 файла Office. js. Так как API JavaScript для Office поддерживает обратную совместимость, последний выпуск по-прежнему будет поддерживать элементы API, представленные ранее в версии 1. Если вам нужно обновить существующий проект, ознакомьтесь со статьей [Обновление версии API JavaScript для Office и файлов схемы манифеста](update-your-javascript-api-for-office-and-manifest-schema-version.md). 

Если вы планируете опубликовать свою надстройку Office из AppSource, необходимо использовать эту ссылку на сеть CDN. Локальные ссылки подходят только для внутренних сценариев, а также сценариев разработки и отладки.

> [!NOTE]
> Чтобы использовать предварительные API-интерфейсы, сослаться на предварительную версию библиотеки API JavaScript для `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`Office в сети CDN:.

## <a name="enabling-intellisense-for-a-typescript-project"></a>Включение IntelliSense для проекта TypeScript

Кроме ссылки на API JavaScript для Office, как описано выше, можно также включить функцию IntelliSense для проекта надстройки TypeScript, используя определения типов из [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js). Для этого выполните следующую команду в командной строке с поддержкой узлов (или в окне Bash Git) из корневого каталога папки проекта. У вас должен быть установлен [Node.js](https://nodejs.org) (содержащий npm).

```command&nbsp;line
npm install --save-dev @types/office-js
```

> [!NOTE]
> Чтобы включить IntelliSense для предварительной версии API, используйте следующие команды в корневой папке проекта, [выполнив следующую](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js-preview) команду: 
>
> `npm install --save-dev @types/office-js-preview`

## <a name="see-also"></a>См. также

- [Общие сведения об интерфейсе API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript для Office](/office/dev/add-ins/reference/javascript-api-for-office)
