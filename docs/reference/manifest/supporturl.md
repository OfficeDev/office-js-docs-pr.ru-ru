# <a name="supporturl-element"></a>Элемент SupportUrl

Указывает URL-адрес страницы, на которой представлены сведения о поддержке надстройки.

## <a name="syntax"></a>Синтаксис

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

## <a name="contained-in"></a>Элемент, в котором содержится

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Может содержать

|  Элемент | Обязательный | Описание  |
|:-----|:-----|:-----|
|  [Переопределение](override.md)   | Нет | Задает параметр для URL-адресов дополнительных языковых стандартов |

## <a name="attributes"></a>Атрибуты

|**Атрибут**|**Тип**|**Обязательный**|**Описание**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|обязательный|Задает значение по умолчанию для этого параметра, представленное для языкового стандарта, который указан с помощью элемента [DefaultLocale](defaultlocale.md).|
