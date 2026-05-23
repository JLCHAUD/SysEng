import React, { useState } from 'react';
import {
  View, Text, StyleSheet, SafeAreaView, TouchableOpacity,
  ActivityIndicator, Alert,
} from 'react-native';
import { Ionicons } from '@expo/vector-icons';
import RecordButton from '../components/RecordButton';
import { transcribeAudio } from '../services/transcription';
import { useStore } from '../store/useStore';
import { Recording } from '../types';
import 'react-native-get-random-values';
import { v4 as uuidv4 } from 'uuid';

interface Props {
  onDone: (rec: Recording) => void;
  onBack: () => void;
}

type Step = 'record' | 'transcribing' | 'done';

export default function RecordScreen({ onDone, onBack }: Props) {
  const { addRecording } = useStore();
  const [step, setStep] = useState<Step>('record');

  async function handleRecordingComplete(uri: string, duration: number) {
    setStep('transcribing');
    try {
      const transcript = await transcribeAudio(uri);
      const now = new Date().toISOString();
      const rec: Recording = {
        id: uuidv4(),
        title: `Enregistrement ${new Date().toLocaleDateString('fr-FR')}`,
        audioUri: uri,
        duration,
        transcript,
        categoryId: null,
        createdAt: now,
        updatedAt: now,
      };
      await addRecording(rec);
      setStep('done');
      onDone(rec);
    } catch (err: any) {
      Alert.alert('Erreur de transcription', err.message ?? 'Une erreur est survenue.');
      setStep('record');
    }
  }

  return (
    <SafeAreaView style={styles.safe}>
      <View style={styles.header}>
        <TouchableOpacity onPress={onBack} hitSlop={12}>
          <Ionicons name="arrow-back" size={24} color="#1e293b" />
        </TouchableOpacity>
        <Text style={styles.title}>Nouvel enregistrement</Text>
        <View style={{ width: 24 }} />
      </View>

      <View style={styles.center}>
        {step === 'record' && (
          <>
            <Text style={styles.hint}>Appuie sur le micro pour commencer</Text>
            <RecordButton onRecordingComplete={handleRecordingComplete} />
          </>
        )}
        {step === 'transcribing' && (
          <>
            <ActivityIndicator size="large" color="#6366f1" />
            <Text style={styles.hint}>Transcription en cours…</Text>
          </>
        )}
      </View>
    </SafeAreaView>
  );
}

const styles = StyleSheet.create({
  safe: { flex: 1, backgroundColor: '#f8fafc' },
  header: { flexDirection: 'row', alignItems: 'center', justifyContent: 'space-between', paddingHorizontal: 20, paddingVertical: 16 },
  title: { fontSize: 18, fontWeight: '700', color: '#1e293b' },
  center: { flex: 1, justifyContent: 'center', alignItems: 'center', gap: 32 },
  hint: { fontSize: 16, color: '#64748b', textAlign: 'center', paddingHorizontal: 40 },
});
