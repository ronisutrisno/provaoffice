import { createIpcTransport, type AgentTransport } from '@prova/agent-core'
import type { AiSettings } from '../../shared/ipc'

/** IPC transport wired to the flows preload bridge (window.flowsApi). */
export function createElectronTransport(getSettings: () => AiSettings): AgentTransport {
  return createIpcTransport<AiSettings>({
    onStream: (listener) => window.flowsApi.onAiStream(listener),
    start: (request) => window.flowsApi.aiStream(request),
    cancel: (requestId) => void window.flowsApi.aiStreamCancel(requestId),
    getSettings,
    unknownErrorText: () => 'An unexpected error occurred.',
    timeoutErrorText: () => 'The model stopped responding (timeout).',
    creditsErrorText: () => 'Credits exhausted.',
    networkErrorText: () => 'Network error reaching the model.',
  })
}
