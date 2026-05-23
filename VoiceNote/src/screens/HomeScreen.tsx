import React, { useEffect, useState } from 'react';
import {
  View, Text, FlatList, TouchableOpacity, StyleSheet,
  SafeAreaView, TextInput, Alert, Modal, Pressable,
} from 'react-native';
import { Ionicons } from '@expo/vector-icons';
import { useStore } from '../store/useStore';
import { Recording, Category } from '../types';
import { format } from 'date-fns';
import { fr } from 'date-fns/locale';

interface Props {
  onOpenRecording: (rec: Recording) => void;
  onNewRecording: () => void;
}

export default function HomeScreen({ onOpenRecording, onNewRecording }: Props) {
  const { recordings, categories, loadRecordings, loadCategories, deleteRecording } = useStore();
  const [search, setSearch] = useState('');
  const [selectedCat, setSelectedCat] = useState<string | null>(null);

  useEffect(() => {
    loadCategories();
    loadRecordings();
  }, []);

  const filtered = recordings.filter((r) => {
    const matchCat = selectedCat ? r.categoryId === selectedCat : true;
    const matchSearch = search
      ? r.title.toLowerCase().includes(search.toLowerCase()) ||
        r.transcript.toLowerCase().includes(search.toLowerCase())
      : true;
    return matchCat && matchSearch;
  });

  function confirmDelete(id: string) {
    Alert.alert('Supprimer', 'Supprimer cet enregistrement ?', [
      { text: 'Annuler', style: 'cancel' },
      { text: 'Supprimer', style: 'destructive', onPress: () => deleteRecording(id) },
    ]);
  }

  function getCategoryColor(catId: string | null) {
    return categories.find((c) => c.id === catId)?.color ?? '#94a3b8';
  }

  function getCategoryName(catId: string | null) {
    return categories.find((c) => c.id === catId)?.name ?? null;
  }

  function fmtDuration(s: number) {
    return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, '0')}`;
  }

  return (
    <SafeAreaView style={styles.safe}>
      <View style={styles.header}>
        <Text style={styles.title}>VoiceNote</Text>
        <TouchableOpacity onPress={onNewRecording} style={styles.fab} activeOpacity={0.8}>
          <Ionicons name="mic" size={22} color="#fff" />
        </TouchableOpacity>
      </View>

      <View style={styles.searchRow}>
        <Ionicons name="search" size={16} color="#94a3b8" style={{ marginLeft: 12 }} />
        <TextInput
          style={styles.search}
          placeholder="Rechercher..."
          placeholderTextColor="#94a3b8"
          value={search}
          onChangeText={setSearch}
        />
      </View>

      {categories.length > 0 && (
        <View style={styles.catRow}>
          <TouchableOpacity
            style={[styles.catChip, !selectedCat && styles.catChipActive]}
            onPress={() => setSelectedCat(null)}
          >
            <Text style={[styles.catText, !selectedCat && styles.catTextActive]}>Tout</Text>
          </TouchableOpacity>
          {categories.map((c) => (
            <TouchableOpacity
              key={c.id}
              style={[styles.catChip, selectedCat === c.id && { backgroundColor: c.color, borderColor: c.color }]}
              onPress={() => setSelectedCat(selectedCat === c.id ? null : c.id)}
            >
              <Text style={[styles.catText, selectedCat === c.id && { color: '#fff' }]}>{c.name}</Text>
            </TouchableOpacity>
          ))}
        </View>
      )}

      <FlatList
        data={filtered}
        keyExtractor={(item) => item.id}
        contentContainerStyle={styles.list}
        ListEmptyComponent={
          <View style={styles.empty}>
            <Ionicons name="mic-off-outline" size={48} color="#cbd5e1" />
            <Text style={styles.emptyText}>Aucun enregistrement</Text>
          </View>
        }
        renderItem={({ item }) => (
          <TouchableOpacity style={styles.card} onPress={() => onOpenRecording(item)} activeOpacity={0.7}>
            <View style={styles.cardLeft}>
              <View style={[styles.dot, { backgroundColor: getCategoryColor(item.categoryId) }]} />
              <View style={{ flex: 1 }}>
                <Text style={styles.cardTitle} numberOfLines={1}>{item.title}</Text>
                <Text style={styles.cardMeta}>
                  {format(new Date(item.createdAt), 'd MMM yyyy · HH:mm', { locale: fr })}
                  {item.duration > 0 && `  ·  ${fmtDuration(item.duration)}`}
                </Text>
                {item.transcript ? (
                  <Text style={styles.cardSnippet} numberOfLines={2}>{item.transcript}</Text>
                ) : (
                  <Text style={styles.cardNoTranscript}>Pas encore transcrit</Text>
                )}
                {getCategoryName(item.categoryId) && (
                  <Text style={[styles.catBadge, { color: getCategoryColor(item.categoryId) }]}>
                    {getCategoryName(item.categoryId)}
                  </Text>
                )}
              </View>
            </View>
            <TouchableOpacity onPress={() => confirmDelete(item.id)} hitSlop={12}>
              <Ionicons name="trash-outline" size={18} color="#cbd5e1" />
            </TouchableOpacity>
          </TouchableOpacity>
        )}
      />
    </SafeAreaView>
  );
}

const styles = StyleSheet.create({
  safe: { flex: 1, backgroundColor: '#f8fafc' },
  header: { flexDirection: 'row', alignItems: 'center', justifyContent: 'space-between', paddingHorizontal: 20, paddingTop: 16, paddingBottom: 8 },
  title: { fontSize: 28, fontWeight: '800', color: '#1e293b' },
  fab: { backgroundColor: '#6366f1', width: 42, height: 42, borderRadius: 21, justifyContent: 'center', alignItems: 'center' },
  searchRow: { flexDirection: 'row', alignItems: 'center', backgroundColor: '#fff', marginHorizontal: 20, marginBottom: 12, borderRadius: 12, borderWidth: 1, borderColor: '#e2e8f0' },
  search: { flex: 1, paddingHorizontal: 10, paddingVertical: 10, fontSize: 15, color: '#1e293b' },
  catRow: { flexDirection: 'row', paddingHorizontal: 20, gap: 8, marginBottom: 12, flexWrap: 'wrap' },
  catChip: { paddingHorizontal: 12, paddingVertical: 5, borderRadius: 16, borderWidth: 1, borderColor: '#e2e8f0', backgroundColor: '#fff' },
  catChipActive: { backgroundColor: '#6366f1', borderColor: '#6366f1' },
  catText: { fontSize: 13, color: '#64748b', fontWeight: '500' },
  catTextActive: { color: '#fff' },
  list: { paddingHorizontal: 20, paddingBottom: 40 },
  card: { backgroundColor: '#fff', borderRadius: 14, padding: 14, marginBottom: 10, flexDirection: 'row', alignItems: 'flex-start', gap: 10, shadowColor: '#000', shadowOpacity: 0.04, shadowRadius: 8, elevation: 2 },
  cardLeft: { flex: 1, flexDirection: 'row', gap: 10 },
  dot: { width: 10, height: 10, borderRadius: 5, marginTop: 5 },
  cardTitle: { fontSize: 16, fontWeight: '600', color: '#1e293b' },
  cardMeta: { fontSize: 12, color: '#94a3b8', marginTop: 2 },
  cardSnippet: { fontSize: 13, color: '#64748b', marginTop: 4, lineHeight: 18 },
  cardNoTranscript: { fontSize: 13, color: '#cbd5e1', marginTop: 4, fontStyle: 'italic' },
  catBadge: { fontSize: 12, fontWeight: '600', marginTop: 6 },
  empty: { alignItems: 'center', marginTop: 80, gap: 12 },
  emptyText: { color: '#94a3b8', fontSize: 16 },
});
