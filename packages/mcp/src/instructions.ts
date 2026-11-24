/**
 * The user's own instructions, as a section appended after the built-in prompt. Browser-safe:
 * the chat app bundles this file, so it must import nothing from Node.
 */

/** Longest instruction text accepted, in characters, after trimming. */
export const MAX_INSTRUCTIONS = 4000;

/** The section to append to a prompt: empty when there is no text, otherwise the trimmed text in a labelled block. */
export function instructionsSection(text: string | undefined): string {
  const trimmed = text?.trim();
  if (!trimmed) return '';
  return [
    "The user's own instructions for this work follow. Where they differ from the guidance above about how to work or reply, follow them. Locked regions and the document's protection are still enforced by the tools.",
    '',
    '<user_instructions>',
    trimmed,
    '</user_instructions>',
  ].join('\n');
}
