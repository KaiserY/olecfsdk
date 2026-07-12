use proc_macro::TokenStream;
use quote::{format_ident, quote};
use syn::{
    Data, DeriveInput, Expr, Fields, GenericArgument, Ident, Lit, Path, PathArguments, Type,
    parse_macro_input, spanned::Spanned,
};

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
    let mut field_idents = Vec::new();
    let mut write_fields = Vec::new();
    let mut size_fields = Vec::new();
    for field in &fields.named {
        let ident = field.ident.as_ref().expect("named field");
        let ty = &field.ty;
        let attrs = sdk_field_attrs(field)?;
        field_idents.push(ident);
        if attrs.remaining {
            require_plain_byte_vec(&attrs, ty, "remaining")?;
            read_fields.push(quote! {
                let sdk_remaining = ::core::convert::TryInto::<usize>::try_into(reader.remaining()?)
                    .map_err(|_| ::olecfsdk::Error::Limit(
                        concat!("remaining length does not fit usize for ", stringify!(#ident)).into(),
                    ))?;
                let #ident = reader.read_vec(sdk_remaining)?;
            });
            write_fields.push(quote! {
                ::std::io::Write::write_all(writer, &self.#ident)?;
            });
            size_fields.push(quote! { self.#ident.len() as u64 });
        } else if let Some(alignment) = attrs.align.as_ref() {
            require_plain_byte_vec(&attrs, ty, "align")?;
            read_fields.push(quote! {
                let #ident = reader.read_alignment(#alignment)?;
            });
            write_fields.push(quote! {
                let sdk_padding = writer.alignment_padding(#alignment)?;
                if self.#ident.len() != sdk_padding {
                    return Err(::olecfsdk::Error::invalid(
                        writer.position()?,
                        concat!("alignment padding mismatch for ", stringify!(#ident)),
                    ));
                }
                ::std::io::Write::write_all(writer, &self.#ident)?;
            });
            size_fields.push(quote! { self.#ident.len() as u64 });
        } else if let Some(condition) = attrs.condition.as_ref() {
            let element = option_element_type(ty).ok_or_else(|| {
                syn::Error::new(field.span(), "condition is only supported on Option fields")
            })?;
            let read_condition = if let Some(mask) = attrs.mask.as_ref() {
                quote! { (#condition & (#mask)) != 0 }
            } else {
                quote! { #condition != 0 }
            };
            let write_condition = if let Some(mask) = attrs.mask.as_ref() {
                quote! { (self.#condition & (#mask)) != 0 }
            } else {
                quote! { self.#condition != 0 }
            };
            let (read_value, write_value, size_value) = type_io_tokens(element);
            read_fields.push(quote! {
                let #ident = if #read_condition { Some(#read_value) } else { None };
            });
            write_fields.push(quote! {
                match (#write_condition, &self.#ident) {
                    (true, Some(value)) => { #write_value }
                    (false, None) => {}
                    _ => return Err(::olecfsdk::Error::invalid(
                        writer.position()?,
                        concat!("condition mismatch for ", stringify!(#ident)),
                    )),
                }
            });
            size_fields.push(quote! {
                self.#ident.as_ref().map_or(0, |value| { #size_value })
            });
        } else if let Some(repr) = attrs.bitflags.as_ref() {
            let repr_ty = Type::Path(syn::TypePath {
                qself: None,
                path: repr.clone().into(),
            });
            let (read_method, write_method, size) = primitive_io(&repr_ty).ok_or_else(|| {
                syn::Error::new(repr.span(), "bitflags repr must be an integer primitive")
            })?;
            read_fields.push(quote! {
                let #ident = <#ty>::from_bits_retain(reader.#read_method()?);
            });
            write_fields.push(quote! {
                writer.#write_method(self.#ident.bits())?;
            });
            size_fields.push(quote! { #size });
        } else if let Some((read_method, write_method, size)) = primitive_io(ty) {
            reject_count(attrs.count.as_ref(), ty)?;
            read_fields.push(quote! { let #ident = reader.#read_method()?; });
            write_fields.push(quote! { writer.#write_method(self.#ident)?; });
            size_fields.push(quote! { #size });
        } else if let Some(len) = byte_array_len(ty) {
            reject_count(attrs.count.as_ref(), ty)?;
            read_fields.push(quote! { let #ident = reader.read_array::<#len>()?; });
            write_fields.push(quote! {
                ::std::io::Write::write_all(writer, &self.#ident)?;
            });
            size_fields.push(quote! { #len });
        } else if let Some((read_method, write_method, element_size, len)) = primitive_array_io(ty)
        {
            reject_count(attrs.count.as_ref(), ty)?;
            read_fields.push(quote! {
                let #ident = {
                    let mut values = [::core::default::Default::default(); #len];
                    for value in &mut values {
                        *value = reader.#read_method()?;
                    }
                    values
                };
            });
            write_fields.push(quote! {
                for value in &self.#ident {
                    writer.#write_method(*value)?;
                }
            });
            size_fields.push(quote! { (#len as u64) * #element_size });
        } else if let Some(element) = vec_element_type(ty) {
            let count = attrs.count.ok_or_else(|| {
                syn::Error::new(field.span(), "Vec fields require #[sdk(count = \"field\")]")
            })?;
            let (read_value, read_size) =
                if let Some((read_method, _, size)) = primitive_io(element) {
                    (quote! { reader.#read_method()? }, size)
                } else {
                    (
                        quote! { <#element as ::olecfsdk::io::SdkRead>::read_from(reader)? },
                        1,
                    )
                };
            read_fields.push(quote! {
                let sdk_count_offset = reader.position()?;
                let sdk_count = ::core::convert::TryInto::<usize>::try_into(#count)
                    .map_err(|_| ::olecfsdk::Error::invalid(
                        sdk_count_offset,
                        concat!("invalid count for ", stringify!(#ident)),
                    ))?;
                reader.ensure_allocation(sdk_count, #read_size as usize)?;
                let mut #ident = ::std::vec::Vec::with_capacity(sdk_count);
                for _ in 0..sdk_count {
                    #ident.push(#read_value);
                }
            });
            let count_field = &count;
            let count_check = quote! {
                let sdk_count_offset = writer.position()?;
                let sdk_expected = ::core::convert::TryInto::<usize>::try_into(self.#count_field)
                    .map_err(|_| ::olecfsdk::Error::invalid(
                        sdk_count_offset,
                        concat!("invalid count for ", stringify!(#ident)),
                    ))?;
                if self.#ident.len() != sdk_expected {
                    return Err(::olecfsdk::Error::invalid(
                        writer.position()?,
                        concat!("count mismatch for ", stringify!(#ident)),
                    ));
                }
            };
            if let Some((_, write_method, size)) = primitive_io(element) {
                write_fields.push(quote! {
                    #count_check
                    for value in &self.#ident {
                        writer.#write_method(*value)?;
                    }
                });
                size_fields.push(quote! { (self.#ident.len() as u64) * #size });
            } else {
                write_fields.push(quote! {
                    #count_check
                    for value in &self.#ident {
                        <#element as ::olecfsdk::io::SdkWrite>::write_to(value, writer)?;
                    }
                });
                size_fields.push(quote! {
                    self.#ident.iter().map(::olecfsdk::io::SdkSize::sdk_size).sum::<u64>()
                });
            }
        } else {
            reject_count(attrs.count.as_ref(), ty)?;
            read_fields.push(quote! {
                let #ident = <#ty as ::olecfsdk::io::SdkRead>::read_from(reader)?;
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
                #(#read_fields)*
                let value = Self { #(#field_idents,)* };
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

#[derive(Default)]
struct FieldAttrs {
    count: Option<Ident>,
    condition: Option<Ident>,
    mask: Option<Expr>,
    align: Option<Expr>,
    remaining: bool,
    bitflags: Option<Ident>,
}

fn sdk_field_attrs(field: &syn::Field) -> syn::Result<FieldAttrs> {
    let mut attrs = FieldAttrs::default();
    for attr in &field.attrs {
        if !attr.path().is_ident("sdk") {
            continue;
        }
        attr.parse_nested_meta(|meta| {
            if meta.path.is_ident("count") || meta.path.is_ident("condition") {
                let target = if meta.path.is_ident("count") {
                    &mut attrs.count
                } else {
                    &mut attrs.condition
                };
                if target.is_some() {
                    return Err(meta.error("duplicate SdkObject field attribute"));
                }
                let value = meta.value()?;
                *target = Some(if value.peek(Lit) {
                    match value.parse::<Lit>()? {
                        Lit::Str(value) => value.parse()?,
                        lit => {
                            return Err(syn::Error::new(lit.span(), "attribute must name a field"));
                        }
                    }
                } else {
                    value.parse()?
                });
                Ok(())
            } else if meta.path.is_ident("bitflags") {
                if attrs.bitflags.is_some() {
                    return Err(meta.error("duplicate bitflags attribute"));
                }
                let value = meta.value()?;
                attrs.bitflags = Some(if value.peek(Lit) {
                    match value.parse::<Lit>()? {
                        Lit::Str(value) => value.parse()?,
                        lit => {
                            return Err(syn::Error::new(
                                lit.span(),
                                "bitflags must name an integer primitive",
                            ));
                        }
                    }
                } else {
                    value.parse()?
                });
                Ok(())
            } else if meta.path.is_ident("mask") || meta.path.is_ident("align") {
                let target = if meta.path.is_ident("mask") {
                    &mut attrs.mask
                } else {
                    &mut attrs.align
                };
                if target.is_some() {
                    return Err(meta.error("duplicate SdkObject field attribute"));
                }
                *target = Some(meta.value()?.parse()?);
                Ok(())
            } else if meta.path.is_ident("remaining") {
                if attrs.remaining {
                    return Err(meta.error("duplicate remaining attribute"));
                }
                attrs.remaining = true;
                Ok(())
            } else {
                Err(meta.error("unsupported SdkObject field attribute"))
            }
        })?;
    }

    let modes = usize::from(attrs.count.is_some())
        + usize::from(attrs.condition.is_some())
        + usize::from(attrs.align.is_some())
        + usize::from(attrs.remaining)
        + usize::from(attrs.bitflags.is_some());
    if modes > 1 {
        return Err(syn::Error::new(
            field.span(),
            "count, condition, align, remaining, and bitflags are mutually exclusive",
        ));
    }
    if attrs.mask.is_some() && attrs.condition.is_none() {
        return Err(syn::Error::new(
            field.span(),
            "mask requires a condition attribute",
        ));
    }
    Ok(attrs)
}

fn require_plain_byte_vec(attrs: &FieldAttrs, ty: &Type, attribute: &str) -> syn::Result<()> {
    let is_byte_vec = vec_element_type(ty)
        .is_some_and(|element| matches!(element, Type::Path(path) if path.path.is_ident("u8")));
    if !is_byte_vec || attrs.count.is_some() || attrs.condition.is_some() {
        return Err(syn::Error::new(
            ty.span(),
            format!("{attribute} is only supported on Vec<u8> fields"),
        ));
    }
    Ok(())
}

fn reject_count(count: Option<&Ident>, ty: &Type) -> syn::Result<()> {
    if count.is_some() {
        Err(syn::Error::new(
            ty.span(),
            "count is only supported on Vec fields",
        ))
    } else {
        Ok(())
    }
}

fn vec_element_type(ty: &Type) -> Option<&Type> {
    let Type::Path(path) = ty else { return None };
    let segment = path.path.segments.last()?;
    if segment.ident != "Vec" {
        return None;
    }
    let PathArguments::AngleBracketed(arguments) = &segment.arguments else {
        return None;
    };
    match arguments.args.first()? {
        GenericArgument::Type(element) => Some(element),
        _ => None,
    }
}

fn option_element_type(ty: &Type) -> Option<&Type> {
    let Type::Path(path) = ty else { return None };
    let segment = path.path.segments.last()?;
    if segment.ident != "Option" {
        return None;
    }
    let PathArguments::AngleBracketed(arguments) = &segment.arguments else {
        return None;
    };
    match arguments.args.first()? {
        GenericArgument::Type(element) => Some(element),
        _ => None,
    }
}

fn type_io_tokens(
    ty: &Type,
) -> (
    proc_macro2::TokenStream,
    proc_macro2::TokenStream,
    proc_macro2::TokenStream,
) {
    if let Some((read_method, write_method, size)) = primitive_io(ty) {
        (
            quote! { reader.#read_method()? },
            quote! { writer.#write_method(*value)?; },
            quote! { #size },
        )
    } else {
        (
            quote! { <#ty as ::olecfsdk::io::SdkRead>::read_from(reader)? },
            quote! { <#ty as ::olecfsdk::io::SdkWrite>::write_to(value, writer)?; },
            quote! { <#ty as ::olecfsdk::io::SdkSize>::sdk_size(value) },
        )
    }
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

fn primitive_array_io(ty: &Type) -> Option<(proc_macro2::Ident, proc_macro2::Ident, u64, &Expr)> {
    let Type::Array(array) = ty else { return None };
    let Type::Path(element) = array.elem.as_ref() else {
        return None;
    };
    let ident = element.path.get_ident()?.to_string();
    let size = primitive_size_by_name(&ident)?;
    Some((
        format_ident!("read_{ident}"),
        format_ident!("write_{ident}"),
        size,
        &array.len,
    ))
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
