import React, { useState, useRef } from 'react';
import { View, TouchableOpacity, Text, StyleSheet, ActivityIndicator } from 'react-native';
import { Audio } from 'expo-av';
import { Ionicons } from '@expo/vector-icons';

interface Props {
  onRecordingComplete: (uri: string, duration: number) => void;
}

export default function RecordButton({ onRecordingComplete }: Props) {
  const [isRecording, setIsRecording] = useState(false);
  const [elapsed, setElapsed] = useState(0);
  const recordingRef = useRef<Audio.Recording | null>(null);
  const intervalRef = useRef<ReturnType<typeof setInterval> | null>(null);

  async function startRecording() {
    await Audio.requestPermissionsAsync();
    await Audio.setAudioModeAsync({ allowsRecordingIOS: true, playsInSilentModeIOS: true });

    const { recording } = await Audio.Recording.createAsync(
      Audio.RecordingOptionsPresets.HIGH_QUALITY
    );
    recordingRef.current = recording;
    setIsRecording(true);
    setElapsed(0);
    intervalRef.current = setInterval(() => setElapsed((e) => e + 1), 1000);
  }

  async function stopRecording() {
    if (!recordingRef.current) return;
    clearInterval(intervalRef.current!);
    setIsRecording(false);

    await recordingRef.current.stopAndUnloadAsync();
    const uri = recordingRef.current.getURI();
    const status = await recordingRef.current.getStatusAsync();
    const duration = Math.round((status as any).durationMillis / 1000);
    recordingRef.current = null;

    if (uri) onRecordingComplete(uri, duration);
  }

  const fmt = (s: number) => `${String(Math.floor(s / 60)).padStart(2, '0')}:${String(s % 60).padStart(2, '0')}`;

  return (
    <View style={styles.container}>
      {isRecording && <Text style={styles.timer}>{fmt(elapsed)}</Text>}
      <TouchableOpacity
        style={[styles.button, isRecording && styles.recording]}
        onPress={isRecording ? stopRecording : startRecording}
        activeOpacity={0.8}
      >
        <Ionicons name={isRecording ? 'stop' : 'mic'} size={36} color="#fff" />
      </TouchableOpacity>
      <Text style={styles.hint}>{isRecording ? 'Appuie pour arrêter' : 'Appuie pour enregistrer'}</Text>
    </View>
  );
}

const styles = StyleSheet.create({
  container: { alignItems: 'center', gap: 12 },
  button: {
    width: 80, height: 80, borderRadius: 40,
    backgroundColor: '#6366f1',
    justifyContent: 'center', alignItems: 'center',
    shadowColor: '#6366f1', shadowOpacity: 0.4, shadowRadius: 12, elevation: 6,
  },
  recording: { backgroundColor: '#ef4444' },
  timer: { fontSize: 22, fontWeight: '600', color: '#ef4444', letterSpacing: 2 },
  hint: { fontSize: 13, color: '#94a3b8' },
});
