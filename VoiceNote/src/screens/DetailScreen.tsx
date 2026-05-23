import React, { useState } from 'react';
import {
  View, Text, StyleSheet, SafeAreaView, TouchableOpacity,
  TextInput, ScrollView, Alert, Modal, Pressable,
} from 'react-native';
import { Ionicons } from '@expo/vector-icons';
import { Audio } from 'expo-av';
import { useStore } from '../store/useStore';
import AddonPanel from '../components/AddonPanel';
import { Recording } from '../types';

interface Props {
  recording: Recording;
  onBack: () => void;
}

const PALETTE = ['#6366f1', '#ec4899', '#f59e0b', '#10b981', '#3b82f6', '#8b5cf6', '#ef4444'];

export default function DetailScreen({ recording: initial, onBack }: Props) {
  const { categories, updateRecording, addCategory } = useStore();
  const [title, setTitle] = useState(initial.title);
  const [transcript, setTranscript] = useState(initial.transcript);
  const [catId, setCatId] = useState(initial.categoryId);
  const [playing, setPlaying] = useState(false);
  const [soundObj, setSoundObj] = useState<Audio.Sound | null>(null);
  const [showCatModal, setShowCatModal] = useState(false);
  const [newCatName, setNewCatName] = useState('');
  const [newCatColor, setNewCatColor] = useState(PALETTE[0]);

  async function save() {
    await updateRecording(initial.id, { title, transcript, categoryId: catId });
    Alert.alert('Sauvegardé', 'Les modifications ont été enregistrées.');
  }

  async function playPause() {
    if (playing && soundObj) {
      await soundObj.stopAsync();
      await soundObj.unloadAsync();
      setSoundObj(null);
      setPlaying(false);
      return;
    }
    const { sound } = await Audio.Sound.createAsync({ uri: initial.audioUri });
    setSoundObj(sound);
    setPlaying(true);
    await sound.playAsync();
    sound.setOnPlaybackStatusUpdate((status) => {
      if ((status as any).didJustFinish) {
        sound.unloadAsync();
        setSoundObj(null);
        setPlaying(false);
      }
    });
  }

  async function createCategory() {
    if (!newCatName.trim()) return;
    const cat = {
      id: Date.now().toString(),
      name: newCatName.trim(),
      color: newCatColor,
      createdAt: new Date().toISOString(),
    };
    await addCategory(cat);
    setCatId(cat.id);
    setNewCatName('');
    setShowCatModal(false);
  }

  const activeCat = categories.find((c) => c.id === catId);

  return (
    <SafeAreaView style={styles.safe}>
      <View style={styles.header}>
        <TouchableOpacity onPress={onBack} hitSlop={12}>
          <Ionicons name="arrow-back" size={24} color="#1e293b" />
        </TouchableOpacity>
        <TouchableOpacity onPress={save} style={styles.saveBtn} activeOpacity={0.8}>
          <Text style={styles.saveTxt}>Sauvegarder</Text>
        </TouchableOpacity>
      </View>

      <ScrollView contentContainerStyle={styles.content} keyboardShouldPersistTaps="handled">
        <TextInput
          style={styles.titleInput}
          value={title}
          onChangeText={setTitle}
          placeholder="Titre"
          placeholderTextColor="#94a3b8"
        />

        {/* Audio player */}
        <TouchableOpacity style={styles.player} onPress={playPause} activeOpacity={0.8}>
          <Ionicons name={playing ? 'stop-circle' : 'play-circle'} size={32} color="#6366f1" />
          <Text style={styles.playerTxt}>{playing ? 'Arrêter la lecture' : 'Écouter l\'enregistrement'}</Text>
        </TouchableOpacity>

        {/* Category */}
        <View style={styles.section}>
          <Text style={styles.sectionLabel}>Catégorie</Text>
          <ScrollView horizontal showsHorizontalScrollIndicator={false} style={{ marginTop: 8 }}>
            <TouchableOpacity
              style={[styles.catChip, !catId && styles.catChipActive]}
              onPress={() => setCatId(null)}
            >
              <Text style={[styles.catTxt, !catId && { color: '#fff' }]}>Aucune</Text>
            </TouchableOpacity>
            {categories.map((c) => (
              <TouchableOpacity
                key={c.id}
                style={[styles.catChip, catId === c.id && { backgroundColor: c.color, borderColor: c.color }]}
                onPress={() => setCatId(c.id)}
              >
                <Text style={[styles.catTxt, catId === c.id && { color: '#fff' }]}>{c.name}</Text>
              </TouchableOpacity>
            ))}
            <TouchableOpacity style={[styles.catChip, styles.newCat]} onPress={() => setShowCatModal(true)}>
              <Ionicons name="add" size={14} color="#6366f1" />
              <Text style={[styles.catTxt, { color: '#6366f1' }]}>Nouvelle</Text>
            </TouchableOpacity>
          </ScrollView>
        </View>

        {/* Transcript editor */}
        <View style={styles.section}>
          <Text style={styles.sectionLabel}>Transcription</Text>
          <TextInput
            style={styles.transcriptInput}
            value={transcript}
            onChangeText={setTranscript}
            multiline
            placeholder="La transcription apparaîtra ici..."
            placeholderTextColor="#94a3b8"
            textAlignVertical="top"
          />
        </View>

        {/* Add-ons */}
        <View style={styles.section}>
          <AddonPanel text={transcript} onApply={(t) => setTranscript(t)} />
        </View>
      </ScrollView>

      {/* New category modal */}
      <Modal visible={showCatModal} transparent animationType="slide">
        <Pressable style={styles.overlay} onPress={() => setShowCatModal(false)} />
        <View style={styles.modal}>
          <Text style={styles.modalTitle}>Nouvelle catégorie</Text>
          <TextInput
            style={styles.modalInput}
            value={newCatName}
            onChangeText={setNewCatName}
            placeholder="Nom de la catégorie"
            placeholderTextColor="#94a3b8"
            autoFocus
          />
          <View style={styles.palette}>
            {PALETTE.map((color) => (
              <TouchableOpacity
                key={color}
                style={[styles.colorDot, { backgroundColor: color }, newCatColor === color && styles.colorDotActive]}
                onPress={() => setNewCatColor(color)}
              />
            ))}
          </View>
          <TouchableOpacity style={styles.createBtn} onPress={createCategory} activeOpacity={0.8}>
            <Text style={styles.createTxt}>Créer</Text>
          </TouchableOpacity>
        </View>
      </Modal>
    </SafeAreaView>
  );
}

const styles = StyleSheet.create({
  safe: { flex: 1, backgroundColor: '#f8fafc' },
  header: { flexDirection: 'row', alignItems: 'center', justifyContent: 'space-between', paddingHorizontal: 20, paddingVertical: 14 },
  saveBtn: { backgroundColor: '#6366f1', paddingHorizontal: 16, paddingVertical: 7, borderRadius: 10 },
  saveTxt: { color: '#fff', fontWeight: '600', fontSize: 14 },
  content: { padding: 20, gap: 20, paddingBottom: 60 },
  titleInput: { fontSize: 22, fontWeight: '700', color: '#1e293b', borderBottomWidth: 1, borderBottomColor: '#e2e8f0', paddingBottom: 8 },
  player: { flexDirection: 'row', alignItems: 'center', gap: 10, backgroundColor: '#eef2ff', padding: 14, borderRadius: 12 },
  playerTxt: { fontSize: 15, color: '#4f46e5', fontWeight: '500' },
  section: { gap: 4 },
  sectionLabel: { fontSize: 12, fontWeight: '600', color: '#64748b', textTransform: 'uppercase', letterSpacing: 0.8 },
  catChip: { marginRight: 8, paddingHorizontal: 14, paddingVertical: 7, borderRadius: 18, borderWidth: 1, borderColor: '#e2e8f0', backgroundColor: '#fff' },
  catChipActive: { backgroundColor: '#6366f1', borderColor: '#6366f1' },
  catTxt: { fontSize: 13, color: '#64748b', fontWeight: '500' },
  newCat: { flexDirection: 'row', alignItems: 'center', gap: 4, borderStyle: 'dashed', borderColor: '#a5b4fc' },
  transcriptInput: { backgroundColor: '#fff', borderRadius: 12, padding: 14, minHeight: 180, fontSize: 15, color: '#1e293b', lineHeight: 22, borderWidth: 1, borderColor: '#e2e8f0' },
  overlay: { ...StyleSheet.absoluteFill, backgroundColor: 'rgba(0,0,0,0.4)' },
  modal: { position: 'absolute', bottom: 0, left: 0, right: 0, backgroundColor: '#fff', borderTopLeftRadius: 24, borderTopRightRadius: 24, padding: 24, gap: 16 },
  modalTitle: { fontSize: 18, fontWeight: '700', color: '#1e293b' },
  modalInput: { borderWidth: 1, borderColor: '#e2e8f0', borderRadius: 10, padding: 12, fontSize: 15, color: '#1e293b' },
  palette: { flexDirection: 'row', gap: 10 },
  colorDot: { width: 28, height: 28, borderRadius: 14 },
  colorDotActive: { borderWidth: 3, borderColor: '#1e293b' },
  createBtn: { backgroundColor: '#6366f1', padding: 14, borderRadius: 12, alignItems: 'center' },
  createTxt: { color: '#fff', fontWeight: '700', fontSize: 15 },
});
