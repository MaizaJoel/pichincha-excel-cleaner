import 'dart:io';
import 'dart:typed_data';
import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import 'package:path_provider/path_provider.dart';
import 'package:open_filex/open_filex.dart';
import 'package:share_plus/share_plus.dart';
import 'package:intl/intl.dart';

import 'engine.dart';
import 'models.dart';

void main() {
  runApp(const PichinchaCleanerApp());
}

class PichinchaCleanerApp extends StatelessWidget {
  const PichinchaCleanerApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Pichincha Excel Cleaner',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(
          seedColor: const Color(0xFF1F4E78),
          primary: const Color(0xFF1F4E78),
          secondary: const Color(0xFF78C257),
        ),
        useMaterial3: true,
        fontFamily: 'Roboto',
      ),
      home: const HomeScreen(),
    );
  }
}

class HomeScreen extends StatefulWidget {
  const HomeScreen({super.key});

  @override
  State<HomeScreen> createState() => _HomeScreenState();
}

class _HomeScreenState extends State<HomeScreen> {
  String? _fileName;
  ProcessResult? _result;
  bool _isLoading = false;
  String? _errorMessage;

  Future<void> _pickAndProcessFile() async {
    setState(() {
      _isLoading = true;
      _errorMessage = null;
    });

    try {
      FilePickerResult? pickResult = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['xlsx'],
        withData: true,
      );

      if (pickResult == null || pickResult.files.isEmpty) {
        setState(() => _isLoading = false);
        return;
      }

      final file = pickResult.files.first;
      Uint8List? bytes = file.bytes;

      if (bytes == null && file.path != null) {
        bytes = await File(file.path!).readAsBytes();
      }

      if (bytes == null) {
        throw Exception('No se pudo leer el contenido del archivo.');
      }

      final processed = PichinchaMobileEngine.processExcelBytes(bytes);

      setState(() {
        _fileName = file.name;
        _result = processed;
        _isLoading = false;
      });

      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(
            content: Text('¡Éxito! Se limpiaron ${processed.records.length} transacciones.'),
            backgroundColor: const Color(0xFF28A745),
          ),
        );
      }
    } catch (e) {
      setState(() {
        _isLoading = false;
        _errorMessage = e.toString();
      });
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(
            content: Text('Error: ${e.toString()}'),
            backgroundColor: Colors.red,
          ),
        );
      }
    }
  }

  Future<void> _saveAndOpenFile() async {
    if (_result == null) return;

    try {
      final dir = await getApplicationDocumentsDirectory();
      final baseName = _fileName != null ? _fileName!.replaceAll('.xlsx', '') : 'movimientos';
      final outPath = '${dir.path}/${baseName}_filtrable.xlsx';

      final outFile = File(outPath);
      await outFile.writeAsBytes(_result!.excelBytes);

      final openResult = await OpenFilex.open(outPath);
      if (openResult.type != ResultType.done) {
        // Fallback to share sheet
        await Share.shareXFiles([XFile(outPath)], text: 'Movimientos Banco Pichincha Limpios');
      }
    } catch (e) {
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('Error al exportar: $e'), backgroundColor: Colors.red),
        );
      }
    }
  }

  Future<void> _shareFile() async {
    if (_result == null) return;

    try {
      final tempDir = await getTemporaryDirectory();
      final baseName = _fileName != null ? _fileName!.replaceAll('.xlsx', '') : 'movimientos';
      final outPath = '${tempDir.path}/${baseName}_filtrable.xlsx';

      final file = File(outPath);
      await file.writeAsBytes(_result!.excelBytes);

      await Share.shareXFiles(
        [XFile(outPath)],
        text: 'Estado de Cuenta Banco Pichincha - Tabla Limpia y Filtrable',
      );
    } catch (e) {
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('Error al compartir: $e'), backgroundColor: Colors.red),
        );
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    final currencyFmt = NumberFormat.currency(symbol: '\$', decimalDigits: 2);
    final dateFmt = DateFormat('dd-MM-yyyy HH:mm');

    return Scaffold(
      appBar: AppBar(
        title: const Text(
          'Banco Pichincha Excel Cleaner',
          style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold, fontSize: 18),
        ),
        backgroundColor: const Color(0xFF1F4E78),
        actions: [
          Container(
            margin: const EdgeInsets.only(right: 12),
            padding: const EdgeInsets.symmetric(horizontal: 8, vertical: 4),
            decoration: BoxDecoration(
              color: const Color(0xFF153654),
              borderRadius: BorderRadius.circular(12),
            ),
            child: const Row(
              children: [
                Icon(Icons.lock, color: Color(0xFF78C257), size: 14),
                SizedBox(width: 4),
                Text('100% Offline', style: TextStyle(color: Color(0xFF78C257), fontSize: 11, fontWeight: FontWeight.bold)),
              ],
            ),
          )
        ],
      ),
      body: SafeArea(
        child: Column(
          children: [
            // Top Action Card
            Container(
              padding: const EdgeInsets.all(16),
              color: Colors.white,
              child: Column(
                children: [
                  Row(
                    children: [
                      Expanded(
                        child: ElevatedButton.icon(
                          onPressed: _isLoading ? null : _pickAndProcessFile,
                          icon: _isLoading
                              ? const SizedBox(width: 18, height: 18, child: CircularProgressIndicator(strokeWidth: 2, color: Colors.white))
                              : const Icon(Icons.file_upload, color: Colors.white),
                          label: Text(
                            _fileName == null ? 'Seleccionar Archivo XLSX' : 'Cambiar Archivo',
                            style: const TextStyle(color: Colors.white, fontWeight: FontWeight.bold),
                          ),
                          style: ElevatedButton.styleFrom(
                            backgroundColor: const Color(0xFF1F4E78),
                            padding: const EdgeInsets.symmetric(vertical: 14),
                            shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
                          ),
                        ),
                      ),
                    ],
                  ),
                  if (_fileName != null) ...[
                    const SizedBox(height: 8),
                    Text(
                      '📄 $_fileName',
                      style: const TextStyle(fontWeight: FontWeight.bold, color: Color(0xFF1F4E78), fontSize: 13),
                    ),
                  ],
                ],
              ),
            ),

            if (_errorMessage != null)
              Container(
                margin: const EdgeInsets.all(12),
                padding: const EdgeInsets.all(12),
                decoration: BoxDecoration(color: Colors.red.shade50, borderRadius: BorderRadius.circular(8)),
                child: Row(
                  children: [
                    const Icon(Icons.error_outline, color: Colors.red),
                    const SizedBox(width: 8),
                    Expanded(child: Text(_errorMessage!, style: const TextStyle(color: Colors.red))),
                  ],
                ),
              ),

            // Main Content Area
            Expanded(
              child: _result == null
                  ? Center(
                      child: Column(
                        mainAxisAlignment: MainAxisAlignment.center,
                        children: [
                          Icon(Icons.table_view_rounded, size: 70, color: Colors.grey.shade400),
                          const SizedBox(height: 12),
                          Text(
                            'Selecciona tu estado de cuenta de Banco Pichincha (.xlsx)',
                            style: TextStyle(color: Colors.grey.shade600, fontSize: 14),
                            textAlign: TextAlign.center,
                          ),
                          const SizedBox(height: 6),
                          Text(
                            'El archivo se limpiará y transformará en una tabla con filtros.',
                            style: TextStyle(color: Colors.grey.shade500, fontSize: 12),
                            textAlign: TextAlign.center,
                          ),
                        ],
                      ),
                    )
                  : Column(
                      children: [
                        // Metrics Cards
                        Padding(
                          padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 8),
                          child: Row(
                            children: [
                              _metricCard('TRANSACCIONES', '${_result!.records.length}', Colors.black87),
                              const SizedBox(width: 8),
                              _metricCard(
                                'INGRESOS',
                                currencyFmt.format(_result!.summaries.first.totalIncome),
                                const Color(0xFF28A745),
                              ),
                              const SizedBox(width: 8),
                              _metricCard(
                                'GASTOS',
                                currencyFmt.format(_result!.summaries.first.totalExpenses),
                                const Color(0xFFDC3545),
                              ),
                            ],
                          ),
                        ),

                        // Preview Table Header
                        Padding(
                          padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 4),
                          child: Row(
                            mainAxisAlignment: MainAxisAlignment.spaceBetween,
                            children: [
                              const Text('Vista previa', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 13, color: Color(0xFF1F4E78))),
                              Text('${_result!.summaries.first.rowsRemoved} filas eliminadas', style: TextStyle(fontSize: 12, color: Colors.grey.shade600)),
                            ],
                          ),
                        ),

                        // Preview List
                        Expanded(
                          child: ListView.separated(
                            padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 6),
                            itemCount: _result!.records.length > 50 ? 50 : _result!.records.length,
                            separatorBuilder: (_, __) => const Divider(height: 1),
                            itemBuilder: (context, idx) {
                              final rec = _result!.records[idx];
                              final isIncome = rec.tipo == 'Crédito';

                              return ListTile(
                                dense: true,
                                contentPadding: const EdgeInsets.symmetric(horizontal: 8, vertical: 2),
                                title: Text(
                                  rec.concepto,
                                  maxLines: 1,
                                  overflow: TextOverflow.ellipsis,
                                  style: const TextStyle(fontWeight: FontWeight.w600, fontSize: 13),
                                ),
                                subtitle: Text(
                                  '${rec.fecha != null ? dateFmt.format(rec.fecha!) : ''} ${rec.documento.isNotEmpty ? '• Doc: ${rec.documento}' : ''}',
                                  style: TextStyle(fontSize: 11, color: Colors.grey.shade600),
                                ),
                                trailing: Text(
                                  currencyFmt.format(rec.monto),
                                  style: TextStyle(
                                    fontWeight: FontWeight.bold,
                                    color: isIncome ? const Color(0xFF28A745) : const Color(0xFFDC3545),
                                    fontSize: 13,
                                  ),
                                ),
                              );
                            },
                          ),
                        ),
                      ],
                    ),
            ),

            // Bottom Action Bar
            if (_result != null)
              Container(
                padding: const EdgeInsets.all(12),
                decoration: BoxDecoration(
                  color: Colors.white,
                  boxShadow: [BoxShadow(color: Colors.black.withOpacity(0.05), blurRadius: 4, offset: const Offset(0, -2))],
                ),
                child: Row(
                  children: [
                    Expanded(
                      child: OutlinedButton.icon(
                        onPressed: _shareFile,
                        icon: const Icon(Icons.share),
                        label: const Text('Compartir'),
                        style: OutlinedButton.styleFrom(
                          padding: const EdgeInsets.symmetric(vertical: 12),
                          shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
                        ),
                      ),
                    ),
                    const SizedBox(width: 8),
                    Expanded(
                      flex: 2,
                      child: ElevatedButton.icon(
                        onPressed: _saveAndOpenFile,
                        icon: const Icon(Icons.table_chart, color: Colors.white),
                        label: const Text('Abrir en Excel', style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold)),
                        style: ElevatedButton.styleFrom(
                          backgroundColor: const Color(0xFF28A745),
                          padding: const EdgeInsets.symmetric(vertical: 12),
                          shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(8)),
                        ),
                      ),
                    ),
                  ],
                ),
              ),
          ],
        ),
      ),
    );
  }

  Widget _metricCard(String title, String value, Color color) {
    return Expanded(
      child: Container(
        padding: const EdgeInsets.all(10),
        decoration: BoxDecoration(
          color: Colors.white,
          borderRadius: BorderRadius.circular(8),
          border: Border.all(color: Colors.grey.shade200),
        ),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Text(title, style: TextStyle(color: Colors.grey.shade600, fontSize: 10, fontWeight: FontWeight.bold)),
            const SizedBox(height: 2),
            FittedBox(
              fit: BoxFit.scaleDown,
              child: Text(value, style: TextStyle(fontWeight: FontWeight.bold, fontSize: 14, color: color)),
            ),
          ],
        ),
      ),
    );
  }
}
