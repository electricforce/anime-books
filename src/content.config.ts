import { defineCollection } from "astro:content";
import { z } from 'astro/zod';
// 1. Importas el loader glob para archivos locales
import { glob } from 'astro/loaders';

const bookS = defineCollection({
    // 2. Ahora usamos 'loader' en lugar de asumir que están en src/content
    loader: glob({ pattern: "**/*.md", base: "src/content/books" }),

    // 3. El esquema se define directamente
    schema: z.object({
        title: z.string(),
        author: z.string(),
        description: z.string(),
        cover_image: z.string(),
        volume: z.number(),
        publishedDate: z.string(), // O z.coerce.date() si quieres objetos Date
        buy: z.object({
            Amazon: z.string().optional().or(z.literal("")),
            BookDepository: z.string().optional().or(z.literal("")),
            Fnac: z.string().optional().or(z.literal("")),
        }),
        slug: z.string(),
    })
});

export const collections = { bookS };