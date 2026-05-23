import Anthropic from '@anthropic-ai/sdk';
import { TextAddon } from '../types';

const client = new Anthropic({ apiKey: process.env.EXPO_PUBLIC_ANTHROPIC_API_KEY ?? '' });

async function claudeTransform(systemPrompt: string, text: string): Promise<string> {
  const message = await client.messages.create({
    model: 'claude-sonnet-4-6',
    max_tokens: 2048,
    system: systemPrompt,
    messages: [{ role: 'user', content: text }],
  });
  const block = message.content[0];
  return block.type === 'text' ? block.text : text;
}

// Built-in add-ons
export const builtinAddons: TextAddon[] = [
  {
    id: 'remove-fillers',
    name: 'Supprimer les remplissages',
    description: 'Enlève les "euh", "donc", "en fait", répétitions et phrases creuses',
    icon: 'cut-outline',
    run: (text) =>
      claudeTransform(
        `Tu es un éditeur de transcriptions. Supprime uniquement les mots de remplissage (euh, bah, donc, voilà, en fait, etc.), les répétitions inutiles et les phrases qui n'apportent aucun sens. Conserve le sens et le style original. Réponds uniquement avec le texte corrigé, sans commentaire.`,
        text
      ),
  },
  {
    id: 'detect-themes',
    name: 'Détecter les thèmes',
    description: 'Identifie les thèmes principaux et les insère en tête du texte',
    icon: 'pricetags-outline',
    run: (text) =>
      claudeTransform(
        `Tu es un analyste de contenu. Identifie 3 à 5 thèmes principaux du texte suivant. Réponds sous ce format exact :\n\n🏷️ Thèmes : thème1 · thème2 · thème3\n\n---\n\n[texte original sans modification]`,
        text
      ),
  },
];

// Registry — consumers can push their own add-ons here at runtime
const addonRegistry: TextAddon[] = [...builtinAddons];

export function registerAddon(addon: TextAddon): void {
  if (addonRegistry.find((a) => a.id === addon.id)) return;
  addonRegistry.push(addon);
}

export function getAddons(): TextAddon[] {
  return addonRegistry;
}

export function getAddon(id: string): TextAddon | undefined {
  return addonRegistry.find((a) => a.id === id);
}
