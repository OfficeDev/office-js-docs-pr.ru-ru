---
title: 'Ссылки на библиотеку API JavaScript для Office '
description: Узнайте, как ссылаться на библиотеку API JavaScript для Office и определения типов в надстройке.
ms.date: 02/18/2021
ms.localizationpriority: medium
ms.openlocfilehash: 38121fe3d3df0a86fef3e2c8e3a58399640f1e2a
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660118"
---
# <a name="referencing-the-office-javascript-api-library"></a>Ссылки на библиотеку API JavaScript для Office 

Библиотека [API JavaScript для Office](../reference/javascript-api-for-office.md) предоставляет интерфейсы API, которые надстройка может использовать для взаимодействия с приложением Office. Самый простой способ ссылки на библиотеку — использовать сеть доставки содержимого (CDN), `<script>` `<head>` добавив следующий тег в раздел HTML-страницы.

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

При первой загрузке надстройка загрузит и кэширует файлы API JavaScript для Office, чтобы убедиться, что она использует последнюю реализацию Office.js и связанных с ней файлов для указанной версии.

> [!IMPORTANT]
> Необходимо ссылаться на API JavaScript для Office из раздела страницы, чтобы обеспечить полную инициализацию API `<head>` перед любыми элементами основного текста.

## <a name="api-versioning-and-backward-compatibility"></a>Управление версиями API и обратная совместимость

В предыдущем фрагменте HTML `/1/` `office.js` перед URL-адресом CDN указывается последний добавочный выпуск в версии 1 Office.js. Так как API JavaScript для Office поддерживает обратную совместимость, последний выпуск будет по-прежнему поддерживать элементы API, которые были представлены ранее в версии 1. Если вам нужно обновить существующий проект, см. статью "Обновление версии [API JavaScript для Office и файлов схемы манифеста"](update-your-javascript-api-for-office-and-manifest-schema-version.md). 

Если вы планируете опубликовать свою надстройку Office из AppSource, необходимо использовать эту ссылку на сеть CDN. Локальные ссылки подходят только для внутренних сценариев, а также сценариев разработки и отладки.

> [!NOTE]
> Чтобы использовать API предварительных версий, требуется указать ссылку на предварительную версию библиотеки API JavaScript для Office в сети CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.

## <a name="enabling-intellisense-for-a-typescript-project"></a>Включение IntelliSense для проекта TypeScript

Помимо ссылки на API JavaScript для Office, как описано выше, вы также можете включить IntelliSense для проекта надстройки TypeScript, используя определения типов из [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js). Для этого выполните следующую команду в системной командной строке с поддержкой узла (или в окне Git bash) из корневой папки проекта. У вас должен быть установлен [Node.js](https://nodejs.org) (содержащий npm).

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a>API предварительной версии

Новые API JavaScript впервые появились в предварительной версии, а затем становятся частью определенного нумерованного набора обязательных элементов после достаточного тестирования и получения отзывов пользователей.

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a>См. также

- [Общие сведения об API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript для Office](../reference/javascript-api-for-office.md)
