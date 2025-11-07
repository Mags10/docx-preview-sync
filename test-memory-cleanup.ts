import { renderSync, renderAsync } from './src/docx-preview';

// Función para simular carga de documento (en un entorno real esto sería un archivo DOCX)
function createMockDocument(): Blob {
    // Crear un documento DOCX mínimo simulado como Blob
    const mockContent = '<?xml version="1.0" encoding="UTF-8"?><document></document>';
    return new Blob([mockContent], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}

// Función para medir uso de memoria (aproximado)
function getMemoryUsage(): number {
    if (typeof performance !== 'undefined' && (performance as any).memory) {
        return (performance as any).memory.usedJSHeapSize;
    }
    return 0;
}

// Función de prueba para verificar limpieza de memoria durante refresh
async function testMemoryCleanup(iterations: number = 10): Promise<void> {
    console.log('🧪 Iniciando prueba de limpieza de memoria durante refresh...');

    // Crear contenedor de prueba
    const bodyContainer = document.createElement('div');
    const styleContainer = document.createElement('div');
    document.body.appendChild(bodyContainer);
    document.body.appendChild(styleContainer);

    const memoryReadings: number[] = [];
    const mockDoc = createMockDocument();

    for (let i = 0; i < iterations; i++) {
        console.log(`📄 Iteración ${i + 1}/${iterations}`);

        try {
            // Renderizar documento
            await renderSync(mockDoc, bodyContainer, styleContainer);

            // Medir memoria después del renderizado
            const memoryAfterRender = getMemoryUsage();
            if (memoryAfterRender > 0) {
                memoryReadings.push(memoryAfterRender);
                console.log(`   Memoria después del renderizado: ${(memoryAfterRender / 1024 / 1024).toFixed(2)} MB`);
            }

            // Pequeña pausa entre iteraciones
            await new Promise(resolve => setTimeout(resolve, 100));

        } catch (error) {
            console.error(`❌ Error en iteración ${i + 1}:`, error);
        }
    }

    // Análisis de resultados
    if (memoryReadings.length > 1) {
        const firstReading = memoryReadings[0];
        const lastReading = memoryReadings[memoryReadings.length - 1];
        const memoryIncrease = lastReading - firstReading;
        const increasePercentage = (memoryIncrease / firstReading) * 100;

        console.log('\n📊 Resultados del análisis de memoria:');
        console.log(`   Lectura inicial: ${(firstReading / 1024 / 1024).toFixed(2)} MB`);
        console.log(`   Lectura final: ${(lastReading / 1024 / 1024).toFixed(2)} MB`);
        console.log(`   Aumento total: ${(memoryIncrease / 1024 / 1024).toFixed(2)} MB (${increasePercentage.toFixed(2)}%)`);

        if (Math.abs(increasePercentage) < 10) { // Menos del 10% de aumento
            console.log('✅ ¡EXCELENTE! La limpieza de memoria está funcionando correctamente.');
            console.log('   El aumento de memoria es mínimo, indicando que no hay fugas significativas.');
        } else if (Math.abs(increasePercentage) < 25) { // Menos del 25% de aumento
            console.log('⚠️  La limpieza de memoria es aceptable, pero podría mejorarse.');
            console.log('   Hay un ligero aumento de memoria que podría indicar pequeñas fugas.');
        } else {
            console.log('❌ POSIBLE PROBLEMA: Aumento significativo de memoria detectado.');
            console.log('   Puede haber fugas de memoria que necesitan ser abordadas.');
        }
    }

    // Limpiar contenedores de prueba
    document.body.removeChild(bodyContainer);
    document.body.removeChild(styleContainer);

    console.log('\n🏁 Prueba completada.');
}

// Ejecutar prueba si estamos en un navegador
if (typeof window !== 'undefined') {
    // Esperar a que el DOM esté listo
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => {
            testMemoryCleanup(5); // Ejecutar con 5 iteraciones para prueba rápida
        });
    } else {
        testMemoryCleanup(5);
    }
}

export { testMemoryCleanup };