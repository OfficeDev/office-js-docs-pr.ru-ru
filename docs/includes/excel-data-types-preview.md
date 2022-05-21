> [!NOTE]
> API типов данных в настоящее время можно использовать только в общедоступной предварительной версии. API предварительной версии могут быть изменены и не предназначены для использования в рабочей среде. Рекомендуется использовать их только в тестовой среде и среде разработки. Не используйте API предварительной версии в рабочей среде или в важных деловых документах.
>
> Чтобы использовать API предварительной версии:
>
> - Необходимо ссылаться на **бета-версию** библиотеки в сети доставки содержимого (CDN) (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). [Файл определения типа](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) для компиляции TypeScript и IntelliSense находится в сети CDN и имеет тип [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Эти типы можно установить с помощью `npm install --save-dev @types/office-js-preview`. Дополнительные сведения см. в файле сведений пакета NPM [@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js).
> - Возможно, вам потребуется присоединиться к [программе предварительной оценки Office](https://insider.office.com), чтобы получить доступ к более поздним сборкам Office.
>
> Чтобы попробовать типы данных в Office для Windows, номер вашей сборки Excel должен быть не ниже 16.0.14626.10000. Чтобы попробовать типы данных в Office для Mac, номер вашей сборки Excel должен быть не ниже 16.55.21102600.