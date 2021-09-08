---
title: Разработка надстроек Office с помощью Angular
description: Используйте Angular для создания надстройки Office в качестве приложения для одной страницы.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: e0d30b7cb2f3d5489f5dae9e257c0cfc115a955e
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936294"
---
# <a name="develop-office-add-ins-with-angular"></a>Разработка надстроек Office с помощью Angular

В этой статье приведены рекомендации по использованию Angular 2 и более поздних версий для создания надстройки Office в виде одностраничного приложения.

> [!NOTE]
> Вы можете поделиться опытом по использованию Angular для создания надстроек Office? Вы можете внести свой вклад в эту статью [в GitHub](https://github.com/OfficeDev/office-js-docs-pr/blob/master/docs/develop/add-ins-with-angular2.md) или предоставить свои отзывы, подав вопрос [в](https://github.com/OfficeDev/office-js-docs-pr/issues) репо.

Пример надстройки Office, созданной на платформе Angular, приведен в статье [Надстройка на основе Angular для проверки стиля в приложении Word](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker).

## <a name="install-the-typescript-type-definitions"></a>Установка определений типов TypeScript

Откройте окно Node.js и введите следующее в командной строке.

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="bootstrapping-must-be-inside-officeinitialize"></a>Начальная загрузка должна определяться в методе Office.initialize

На любой странице, на которую Office, Word или Excel API JavaScript, код должен сначала назначить метод `Office.initialize` свойству. (Если у вас нет кода инициализации, тело метода может быть просто пустым " символы", но вы не должны оставлять свойство `{}` `Office.initialize` неопределенным. Дополнительные сведения см. в [материале Initialize your Office надстройки.)](initialize-add-in.md) Office вызывает этот метод сразу после инициализации Office JavaScript.

**Вызов кода начальной загрузки на основе Angular необходимо задать в методе, который назначен `Office.initialize`**, чтобы сначала выполнялась инициализация библиотек JavaScript для Office. Вот простой пример, в котором показано, как это сделать. Этот код должен находиться в файле main.ts проекта.

```js
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { AppModule } from './app.module';

Office.initialize = function () {
  const platform = platformBrowserDynamic();
  platform.bootstrapModule(AppModule);
};
```

## <a name="use-the-hash-location-strategy-in-the-angular-application"></a>Используйте стратегию навигации с помощью хэша в приложении на основе Angular

Переход между маршрутами в приложении может не выполняться, если не задать стратегию навигации с помощью хэша. Это можно сделать одним из двух способов. Способ первый: указать поставщика стратегии навигации в модуле приложения, как показано в приведенном ниже примере. (Это для файла app.module.ts.)

```js
import { LocationStrategy, HashLocationStrategy } from '@angular/common';
// Other imports suppressed for brevity

@NgModule({
  providers: [
    { provide: LocationStrategy, useClass: HashLocationStrategy },
    // Other providers suppressed
  ],
  // Other module properties suppressed
})
export class AppModule { }
```

Если маршруты определены в отдельном модуле маршрутизации, можно задать стратегию навигации с помощью хэша иначе. В TS-файле модуля маршрутизации передайте объект конфигурации в функцию `forRoot`, которая определяет стратегию. Ниже приведен код в качестве примера.

```js
import { RouterModule, Routes } from '@angular/router';
// Other imports suppressed for brevity

const routes: Routes = // route definitions go here

@NgModule({
  imports: [RouterModule.forRoot(routes, { useHash: true })],
  exports: [RouterModule]
})
export class AppRoutingModule { }
```

## <a name="use-the-office-dialog-api-with-angular"></a>Используйте API Office диалогов с Angular

Надстройка dialog AP в Office позволяет вышей надстройке открывать страницы в немодальном диалоговом окне, с помощью которой можно обмениваться информацией с главной страницей, которая обычно находится в панели задач.

Метод [displayDialogAsync](/javascript/api/office/office.ui) принимает параметр, определяющий URL-адрес страницы, которую нужно открыть в диалоговом окне. В вашей надстройке может быть отдельная HTML-страница (отличающаяся от базовой) для передачи в этот параметр, или же вы можете передать URL-адрес маршрута в программе Angular.

Важно помнить, что в случае передачи маршрута диалоговое окно создает новое окно с собственным контекстом выполнения. Базовая страница со всем ее кодом инициализации и начальной загрузки запускается снова в этом новом контексте, а возможным переменным присваиваются первоначальные значения в диалоговом окне. Такой способ приводит к запуску второго экземпляра одностраничного приложения в диалоговом окне. Код, меняющий переменные в диалоговом окне, не меняет версию области задач этих переменных. Кроме того, диалоговое окно имеет собственное хранилище сеансов (свойство [Window.sessionStorage),](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) которое не доступно из кода в области задач.  

## <a name="trigger-the-ui-update"></a>Запуск обновления пользовательского интерфейса

В приложении Angular пользовательский интерфейс иногда не обновляется. Это происходит потому, что эта часть кода выполняется вне зоны Angular. Чтобы решить эту проблему, поместите код в зону, как показано в приведенном ниже примере.

```js
import { NgZone } from '@angular/core';

export class MyComponent {
  constructor(private zone: NgZone) { }

  myFunction() {
    this.zone.run(() => {
      // the codes that need update the UI
    });
  }
}
```

## <a name="use-observable"></a>Использование Observable

Angular использует библиотеку RxJS, в которой предусмотрены объекты `Observable` и `Observer` для реализации асинхронной обработки. Из этого раздела вы узнаете, как использовать `Observables`. Более подробную информацию см. в официальной документации по [RxJS](https://rxjs-dev.firebaseapp.com/).

Объект `Observable` отчасти похож на объект `Promise`: он возвращается сразу же после асинхронного вызова, но для его разрешения может потребоваться некоторое время. Но если `Promise` — это единственное значение (которое может быть объектом массивов), то `Observable` — это массив объектов (возможно, только с одним элементом). Благодаря этому код может вызывать такие [методы массива](https://www.w3schools.com/jsref/jsref_obj_array.asp), как `concat`, `map` и `filter`, для объектов `Observable`.

### <a name="push-instead-of-pull"></a>Push instead of pull

Ваш код извлекает объекты `Promise`, назначая их переменным, тогда как объекты `Observable` рассылают свои значения объектам, которые *подписаны* на `Observable`. Подписчики — объекты `Observer`. Преимущество подхода, предусматривающего подобную рассылку, состоит в том, что позже можно добавлять в массив `Observable` новые элементы. При добавлении нового элемента все объекты `Observer`, подписанные на `Observable`, получают уведомление.

Объект `Observer` настроен на обработку каждого нового объекта (именуемого "следующим") с помощью функции. (Он также настроен на то, чтобы отвечать на ошибку и уведомление о завершении. См. пример в следующем разделе.) По этой причине объекты `Observable` можно использовать в более широком диапазоне сценариев, чем объекты `Promise`. Например, в дополнение к возврату `Observable` при вызове AJAX (этим способом можно вернуть также `Promise`), объект `Observable` можно возвращать из обработчика событий, например обработчика событий изменения для текстового поля. Каждый раз, когда пользователь вводит текст в поле, все подписанные объекты `Observer` немедленно реагируют, используя последний текст или текущее состояние приложения в качестве вводных данных.

### <a name="wait-until-all-asynchronous-calls-have-completed"></a>Подождите, пока все асинхронные вызовы не будут завершены

Чтобы обратный вызов выполнялся только при условии разрешения каждого элемента из набора объектов `Promise`, используйте метод `Promise.all()`.

```js
myPromise.all([x, y, z]).then(
  // TODO: Callback logic goes here
)
```

Чтобы сделать то же самое с объектом `Observable`, используйте метод [Observable.forkJoin()](https://github.com/Reactive-Extensions/RxJS/blob/master/doc/api/core/operators/forkjoin.md).  

```js
const source = Observable.forkJoin([x, y, z]);

const subscription = source.subscribe(
  x => {
    // TODO: Callback logic goes here
  },
  err => console.log('Error: ' + err),
  () => console.log('Completed')
);
```

## <a name="compile-the-angular-application-using-the-ahead-of-time-aot-compiler"></a>Компиляция приложения Angular с помощью компилятора Ahead-of-Time (AOT)

Производительность приложения — одна из наиболее важных составляющих взаимодействия с пользователем. Компилятор Ahead-of-Time (AOT) позволяет оптимизировать приложение Angular, чтобы компилировать приложение во время сборки. Он полностью преобразовывает исходный код (шаблоны HTML и TypeScript) в эффективный код JavaScript. Если приложение скомпилировано с помощью компилятора AOT, в среде выполнения не будет происходить дополнительная компиляция, что ускорит обработку и выполнение асинхронных запросов для шаблонов HTML. Кроме того, уменьшится общий размер приложения, так как компилятор Angular не придется включать в распространяемый файл приложения.

Чтобы использовать компилятор AOT, добавьте `--aot` к команде `ng build` или `ng serve`:

```command&nbsp;line
ng build --aot
ng serve --aot
```

> [!NOTE]
> Дополнительные сведения о компиляторе Ahead-of-Time (AOT) приложения Angular см. в [официальном руководстве](https://angular.io/guide/aot-compiler).

## <a name="support-internet-explorer-if-youre-dynamically-loading-officejs"></a>Поддержка Internet Explorer при динамической загрузке Office.js

На основе Windows версии и Office настольного клиента, в котором работает надстройка, надстройка может использовать Internet Explorer 11. (Дополнительные сведения см. в [браузерах, используемых Office надстройки.)](../concepts/browsers-used-by-office-web-add-ins.md) Angular зависит от нескольких API, но эти API не работают в времени запуска IE, встроенном в Windows `window.history` настольных клиентов. Если эти API не работают, надстройка может не работать должным образом, например, она может загрузить пустую области задач. Чтобы смягчить это, Office.js обнуляет эти API. Однако при динамической загрузке Office.js AngularJS может загрузиться до Office.js. В этом случае необходимо отключить API, добавив следующий код на страницу `window.history` **index.html.**

```js
<script type="text/javascript">window.history.replaceState=null;window.history.pushState=null;</script>
```
