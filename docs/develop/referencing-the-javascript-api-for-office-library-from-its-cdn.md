---
title: 'Ссылки на библиотеку API JavaScript для Office '
description: Узнайте, как ссылаться на библиотеку API JavaScript для Office и определение типов в надстройке.
ms.date: 06/23/2020
localization_priority: Normal
ms.openlocfilehash: 64dd08329b7bbc8c249bd270a431b6cbe93ec52c
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293186"
---
# <a name="referencing-the-office-javascript-api-library"></a>Ссылки на библиотеку API JavaScript для Office 

Библиотека [API JavaScript для Office](../reference/javascript-api-for-office.md) предоставляет API, которые надстройка может использовать для взаимодействия с приложением Office. Самый простой способ добавить ссылку на библиотеку — использовать сеть доставки содержимого (CDN), добавив следующий `<script>` тег в `<head>` раздел страницы HTML:  

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

Это приведет к скачиванию и кэшированию файлов API JavaScript для Office при первом запуске надстройки, чтобы убедиться, что она использует самую актуальную реализацию Office.js и связанные с ней файлы для указанной версии.

> [!IMPORTANT]
> Необходимо ссылаться на API JavaScript для Office из `<head>` раздела страницы, чтобы убедиться, что API полностью инициализирован до элементов основного текста. Приложения Office требуют, чтобы надстройки инициализирулись в течение 5 секунд после активации. Если надстройка не активируется в этом пороговом значении, она будет объявлена без ответа, а пользователю будет выведено сообщение об ошибке.

## <a name="api-versioning-and-backward-compatibility"></a>Управление версиями и обратная совместимость API

В предыдущем фрагменте кода HTML ( `/1/` перед в `office.js` URL-адресе CDN) указывает последний добавочный выпуск в версии 1 Office.js. Так как API JavaScript для Office поддерживает обратную совместимость, последний выпуск по-прежнему будет поддерживать элементы API, представленные ранее в версии 1. Если вам нужно обновить существующий проект, ознакомьтесь со статьей [Обновление версии API JavaScript для Office и файлов схемы манифеста](update-your-javascript-api-for-office-and-manifest-schema-version.md). 

Если вы планируете опубликовать свою надстройку Office из AppSource, необходимо использовать эту ссылку на сеть CDN. Локальные ссылки подходят только для внутренних сценариев, а также сценариев разработки и отладки.

> [!NOTE]
> Чтобы использовать API предварительных версий, требуется указать ссылку на предварительную версию библиотеки API JavaScript для Office в сети CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.

## <a name="enabling-intellisense-for-a-typescript-project"></a>Включение IntelliSense для проекта TypeScript

Кроме ссылки на API JavaScript для Office, как описано выше, можно также включить функцию IntelliSense для проекта надстройки TypeScript, используя определения типов из [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js). Для этого выполните следующую команду в командной строке с поддержкой узлов (или в окне Bash Git) из корневого каталога папки проекта. У вас должен быть установлен [Node.js](https://nodejs.org) (содержащий npm).

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a>Предварительный просмотр API

Новые API JavaScript впервые появляются в "предварительной версии", а затем становятся частью определенного нумерованного набора требований после выполнения достаточного тестирования и необходимости отзыва пользователей.

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a>См. также

- [Общие сведения об API JavaScript для Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript для Office](../reference/javascript-api-for-office.md)
