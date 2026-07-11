use proc_macro::TokenStream;
use quote::{format_ident, quote};
use syn::{Data, DeriveInput, Expr, Fields, Lit, Path, Type, parse_macro_input, spanned::Spanned};

#[proc_macro_derive(SdkObject, attributes(sdk))]
pub fn sdk_object(input: TokenStream) -> TokenStream {
    let input = parse_macro_input!(input as DeriveInput);
    expand_sdk_object(&input)
        .unwrap_or_else(syn::Error::into_compile_error)
        .into()
}

#[proc_macro_derive(SdkEnum, attributes(sdk))]
pub fn sdk_enum(input: TokenStream) -> TokenStream {
    let input = parse_macro_input!(input as DeriveInput);
    expand_sdk_enum(&input)
        .unwrap_or_else(syn::Error::into_compile_error)
        .into()
}

fn expand_sdk_object(input: &DeriveInput) -> syn::Result<proc_macro2::TokenStream> {
    let name = &input.ident;
    let validate = sdk_validate(input)?;
    let fields = match &input.data {
        Data::Struct(data) => match &data.fields {
            Fields::Named(fields) => fields,
            _ => {
                return Err(syn::Error::new(
                    data.fields.span(),
                    "SdkObject requires a struct with named fields",
                ));
            }
        },
        _ => {
            return Err(syn::Error::new(
                input.span(),
                "SdkObject can only be derived for structs",
            ));
        }
    };

    let mut read_fields = Vec::new();
    let mut write_fields = Vec::new();
    let mut size_fields = Vec::new();
    for field in &fields.named {
        let ident = field.ident.as_ref().expect("named field");
        let ty = &field.ty;
        if let Some((read_method, write_method, size)) = primitive_io(ty) {
            read_fields.push(quote! { #ident: reader.#read_method()? });
            write_fields.push(quote! { writer.#write_method(self.#ident)?; });
            size_fields.push(quote! { #size });
        } else if let Some(len) = byte_array_len(ty) {
            read_fields.push(quote! { #ident: reader.read_array::<#len>()? });
            write_fields.push(quote! { writer.write_all(&self.#ident)?; });
            size_fields.push(quote! { #len });
        } else {
            read_fields.push(quote! {
                #ident: <#ty as ::olecfsdk::io::SdkRead>::read_from(reader)?
            });
            write_fields.push(quote! {
                <#ty as ::olecfsdk::io::SdkWrite>::write_to(&self.#ident, writer)?;
            });
            size_fields.push(quote! {
                <#ty as ::olecfsdk::io::SdkSize>::sdk_size(&self.#ident)
            });
        }
    }

    let validate_read = validate.as_ref().map(|path| quote! { #path(&value)?; });
    let validate_write = validate.as_ref().map(|path| quote! { #path(self)?; });

    Ok(quote! {
        impl ::olecfsdk::io::SdkRead for #name {
            fn read_from<R: ::std::io::Read + ::std::io::Seek>(
                reader: &mut ::olecfsdk::io::Reader<R>,
            ) -> ::olecfsdk::Result<Self> {
                let value = Self { #(#read_fields,)* };
                #validate_read
                Ok(value)
            }
        }

        impl ::olecfsdk::io::SdkWrite for #name {
            fn write_to<W: ::std::io::Write + ::std::io::Seek>(
                &self,
                writer: &mut ::olecfsdk::io::Writer<W>,
            ) -> ::olecfsdk::Result<()> {
                #validate_write
                #(#write_fields)*
                Ok(())
            }
        }

        impl ::olecfsdk::io::SdkSize for #name {
            fn sdk_size(&self) -> u64 {
                0 #(+ #size_fields as u64)*
            }
        }
    })
}

fn sdk_validate(input: &DeriveInput) -> syn::Result<Option<Path>> {
    let mut validate = None;
    for attr in &input.attrs {
        if !attr.path().is_ident("sdk") {
            continue;
        }
        attr.parse_nested_meta(|meta| {
            if meta.path.is_ident("validate") {
                let value = meta.value()?;
                if value.peek(Lit) {
                    match value.parse::<Lit>()? {
                        Lit::Str(value) => validate = Some(value.parse()?),
                        lit => return Err(syn::Error::new(lit.span(), "validate must be a path")),
                    }
                } else {
                    validate = Some(value.parse()?);
                }
                Ok(())
            } else {
                Err(meta.error("unsupported SdkObject attribute"))
            }
        })?;
    }
    Ok(validate)
}

fn expand_sdk_enum(input: &DeriveInput) -> syn::Result<proc_macro2::TokenStream> {
    let name = &input.ident;
    let repr = sdk_repr(input)?;
    let read_method = format_ident!("read_{repr}");
    let write_method = format_ident!("write_{repr}");
    let size = primitive_size_by_name(&repr).ok_or_else(|| {
        syn::Error::new(input.span(), "SdkEnum repr must be an integer primitive")
    })?;
    let repr_ident = format_ident!("{repr}");
    let variants = match &input.data {
        Data::Enum(data) => &data.variants,
        _ => return Err(syn::Error::new(input.span(), "SdkEnum requires an enum")),
    };

    let mut from_arms = Vec::new();
    let mut raw_arms = Vec::new();
    for variant in variants {
        if !matches!(variant.fields, Fields::Unit) {
            return Err(syn::Error::new(
                variant.span(),
                "SdkEnum only supports fieldless variants",
            ));
        }
        let ident = &variant.ident;
        let (_, value) = variant.discriminant.as_ref().ok_or_else(|| {
            syn::Error::new(
                variant.span(),
                "SdkEnum variants need explicit discriminants",
            )
        })?;
        from_arms.push(quote! { value if value == (#value as #repr_ident) => Some(Self::#ident) });
        raw_arms.push(quote! { Self::#ident => #value as #repr_ident });
    }

    Ok(quote! {
        impl ::olecfsdk::io::SdkEnumValue for #name {
            type Repr = #repr_ident;
            fn from_raw(value: Self::Repr) -> Option<Self> {
                match value { #(#from_arms,)* _ => None }
            }
            fn raw(self) -> Self::Repr {
                match self { #(#raw_arms,)* }
            }
        }

        impl ::olecfsdk::io::SdkRead for #name {
            fn read_from<R: ::std::io::Read + ::std::io::Seek>(
                reader: &mut ::olecfsdk::io::Reader<R>,
            ) -> ::olecfsdk::Result<Self> {
                let offset = reader.position()?;
                let raw = reader.#read_method()?;
                <Self as ::olecfsdk::io::SdkEnumValue>::from_raw(raw).ok_or_else(|| {
                    ::olecfsdk::Error::invalid(offset, format!(
                        "invalid {} value: {}", stringify!(#name), raw
                    ))
                })
            }
        }

        impl ::olecfsdk::io::SdkWrite for #name {
            fn write_to<W: ::std::io::Write + ::std::io::Seek>(
                &self,
                writer: &mut ::olecfsdk::io::Writer<W>,
            ) -> ::olecfsdk::Result<()> {
                writer.#write_method(<Self as ::olecfsdk::io::SdkEnumValue>::raw(*self))
            }
        }

        impl ::olecfsdk::io::SdkSize for #name {
            fn sdk_size(&self) -> u64 { #size }
        }
    })
}

fn sdk_repr(input: &DeriveInput) -> syn::Result<String> {
    let mut repr = None;
    for attr in &input.attrs {
        if !attr.path().is_ident("sdk") {
            continue;
        }
        attr.parse_nested_meta(|meta| {
            if meta.path.is_ident("repr") {
                let value = meta.value()?;
                if value.peek(Lit) {
                    match value.parse::<Lit>()? {
                        Lit::Str(value) => repr = Some(value.value()),
                        lit => {
                            return Err(syn::Error::new(
                                lit.span(),
                                "repr must be an integer primitive",
                            ));
                        }
                    }
                } else {
                    let path: Path = value.parse()?;
                    repr = path.get_ident().map(ToString::to_string);
                }
                Ok(())
            } else {
                Err(meta.error("unsupported SdkEnum attribute"))
            }
        })?;
    }
    repr.ok_or_else(|| syn::Error::new(input.span(), "SdkEnum requires #[sdk(repr = \"u16\")]"))
}

fn primitive_io(ty: &Type) -> Option<(proc_macro2::Ident, proc_macro2::Ident, u64)> {
    let Type::Path(path) = ty else { return None };
    let ident = path.path.get_ident()?.to_string();
    let size = primitive_size_by_name(&ident)?;
    Some((
        format_ident!("read_{ident}"),
        format_ident!("write_{ident}"),
        size,
    ))
}

fn byte_array_len(ty: &Type) -> Option<&Expr> {
    let Type::Array(array) = ty else { return None };
    let Type::Path(element) = array.elem.as_ref() else {
        return None;
    };
    (element.path.is_ident("u8")).then_some(&array.len)
}

fn primitive_size_by_name(name: &str) -> Option<u64> {
    match name {
        "u8" | "i8" => Some(1),
        "u16" | "i16" => Some(2),
        "u32" | "i32" | "f32" => Some(4),
        "u64" | "i64" | "f64" => Some(8),
        _ => None,
    }
}
