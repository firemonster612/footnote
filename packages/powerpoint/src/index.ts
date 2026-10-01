import type { HostModule, ToolEnv } from "@footnote/core/contracts";
import { createChatRegistry, turnHasChanges } from "./chats.ts";
import { diffDecks, formatDeckState } from "./deckState.ts";
import { powerpointOps } from "./ops/index.ts";
import type { DeckState } from "./ops/types.ts";
import { powerpointSkills } from "./skills/index.ts";
import { powerpointSystemPrompt } from "./systemPrompt.ts";
import { createPowerPointTools } from "./tools/index.ts";
import { undoLastTurn, undoTurns } from "./undo.ts";
import { createWriteGuard } from "./writeGuard.ts";

/** A PowerPoint HostModule with its own chat state. Shells use the shared `powerpointModule`; tests make their own. */
export function createPowerPointModule(): HostModule {
  const chats = createChatRegistry();

  async function getContextBlock(env: ToolEnv, chatId: string): Promise<string> {
    chats.bind(env, chatId);
    const chat = chats.get(chatId);
    let state: DeckState;
    try {
      state = await env.host.call<DeckState>("get_deck_state");
    } catch (error) {
      const reason = error instanceof Error ? error.message : String(error);
      return `<deck_state>\nUnavailable: ${reason}\n</deck_state>`;
    }
    const { documentName } = await env.host.status();
    const themeKey = JSON.stringify(state.theme);
    const block = formatDeckState({
      state,
      ...(documentName && { documentName }),
      ...(chat.lastOutline && { changes: diffDecks(chat.lastOutline, state.deck.slides) }),
      includeTheme: themeKey !== chat.lastThemeKey,
      staleSlideIds: state.deck.slides
        .filter((slide) => {
          const seen = chat.observed.get(slide.id);
          return seen !== undefined && seen !== slide.fingerprint;
        })
        .map((slide) => slide.id),
    });
    chat.lastOutline = state.deck.slides;
    chat.lastThemeKey = themeKey;
    return block;
  }

  return {
    kind: "powerpoint",
    ops: powerpointOps,
    systemPrompt: powerpointSystemPrompt,
    createTools: (env) =>
      createPowerPointTools(
        env,
        createWriteGuard(env.host, () => chats.forEnv(env)),
      ),
    getContextBlock,
    undo: {
      beginTurn: chats.beginTurn,
      canUndo: (chatId) => chats.get(chatId).turns.some(turnHasChanges),
      undoLastTurn: (env, chatId) => undoLastTurn(env.host, chats.get(chatId)),
      undoTurns: (env, chatId, turnIds) => undoTurns(env.host, chats.get(chatId), turnIds),
    },
    skills: powerpointSkills,
  };
}

export const powerpointModule: HostModule = createPowerPointModule();
