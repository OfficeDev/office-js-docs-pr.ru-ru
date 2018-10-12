# <a name="onenote-javascript-api-requirement-sets"></a>Наборы требований API JavaScript для OneNote

Наборы требований — это именованные группы требований API. С помощью наборов требований, указанных в манифесте, или проверки в среде выполнения надстройки Office определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

В приведенной ниже таблице перечислены наборы требований для OneNote, ведущие приложения Office, которые их поддерживают, а также версии сборок или даты выхода.

|  Набор требований  |  Office Online | 
|:-----|:-----|
| OneNoteApi 1.1  | Сентябрь 2016 г. |  

## <a name="office-common-api-requirement-sets"></a>Стандартные наборы требований API для Office

Сведения о стандартных наборах требований API см. в статье [Стандартные наборы требований API для Office](office-add-in-requirement-sets.md).

## <a name="onenote-javascript-api-11"></a>API JavaScript для OneNote 1.1 

API JavaScript для OneNote 1.1 — первая версия этого API. Подробнее об API см. [Общие сведения о программировании API JavaScript для OneNote](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).

## <a name="runtime-requirement-support-check"></a>Проверка поддержки требований в среде выполнения

Во время выполнения кода надстройки могут проверять, поддерживает ли ведущее приложение набор требований API, выполняя следующую проверку: 

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', 1.1) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## <a name="manifest-based-requirement-support-check"></a>Проверка поддержки требований в манифесте

Используйте элемент Requirements в манифесте надстройки, чтобы указать ключевые наборы требований или элементы API, которые должна использовать надстройка. Если платформа или ведущее приложение Office не поддерживает наборы требований или элементы API, указанные в элементе Requirements, надстройка не будет работать в этом ведущем приложении или на этой платформе, а также не будет отображаться в разделе «Мои надстройки».

Ниже показана надстройка, которая загружается во всех ведущих приложениях Office, поддерживающих набор требований OneNoteApi версии 1.1.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a>См. также

- [Версии Office и наборы требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и требований API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
