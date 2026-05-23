import OpenAI from 'openai';
import * as FileSystem from 'expo-file-system';

// API key injected via env — set EXPO_PUBLIC_OPENAI_API_KEY in .env
const client = new OpenAI({ apiKey: process.env.EXPO_PUBLIC_OPENAI_API_KEY ?? '' });

export async function transcribeAudio(audioUri: string): Promise<string> {
  const fileInfo = await FileSystem.getInfoAsync(audioUri);
  if (!fileInfo.exists) throw new Error('Audio file not found');

  // Read file as base64 then create a Blob for the API
  const base64 = await FileSystem.readAsStringAsync(audioUri, {
    encoding: FileSystem.EncodingType.Base64,
  });
  const binaryStr = atob(base64);
  const bytes = new Uint8Array(binaryStr.length);
  for (let i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i);
  const blob = new Blob([bytes], { type: 'audio/m4a' });
  const file = new File([blob], 'recording.m4a', { type: 'audio/m4a' });

  const response = await client.audio.transcriptions.create({
    model: 'whisper-1',
    file,
    language: 'fr',
  });

  return response.text;
}
