/**
 * Copyright 2023-present DreamNum Co., Ltd.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import type { ICalculateFormulaService, ITraceSample } from '@univerjs/engine-formula';
import type { FFormula } from '@univerjs/engine-formula/facade';
import { mkdirSync, writeFileSync } from 'node:fs';
import { Session } from 'node:inspector/promises';
import { join } from 'node:path';
import process from 'node:process';

export interface IProfileResult {
    trace: ITraceSample[];
    peakHeapMB: number;
    finalHeapMB: number;
    wallMs: number;
    cpuProfilePath?: string;
    heapProfilePath?: string;
}

const PROFILE_DIR = join(process.cwd(), '.formula-profiles');

function ensureProfileDir() {
    mkdirSync(PROFILE_DIR, { recursive: true });
}

function profileFilename(label: string, ext: string): string {
    const ts = new Date().toISOString().replace(/[:.]/g, '-');
    return `${label}-${ts}.${ext}`;
}

export function printTrace(trace: ITraceSample[]) {
    if (trace.length === 0) return;

    const peakHeap = Math.max(...trace.map((s) => s.heapMB));
    const final = trace[trace.length - 1];

    console.warn(`\n--- Memory Trace (${trace.length} samples) ---`);
    console.warn(`formulas: ${final.total}`);
    console.warn(`peak heap: ${peakHeap} MB | final heap: ${final.heapMB} MB | wall: ${final.ms} ms`);
    console.warn('progress:');
    for (const s of trace) {
        const pct = Math.round((s.i / s.total) * 100);
        const bar = '█'.repeat(Math.round(pct / 2.5));
        console.warn(`  ${String(pct).padStart(3)}% │ ${String(s.heapMB).padStart(5)} MB │ ${String(s.ms).padStart(7)} ms │${bar}`);
    }
    console.warn('---\n');
}

export async function profileCalculation(
    formulaEngine: FFormula,
    calcService: ICalculateFormulaService,
    options: {
        label?: string;
        timeoutMs?: number;
        cpu?: boolean;
        heap?: boolean;
    } = {}
): Promise<IProfileResult> {
    const { label = 'calc', timeoutMs = 600_000, cpu = false, heap = false } = options;

    let session: Session | null = null;

    if (cpu || heap) {
        session = new Session();
        session.connect();
    }

    if (cpu) {
        await session!.post('Profiler.enable');
        await session!.post('Profiler.start');
    }

    if (heap) {
        await session!.post('HeapProfiler.enable');
        await session!.post('HeapProfiler.startSampling');
    }

    formulaEngine.executeCalculation();
    await formulaEngine.onCalculationEnd(timeoutMs);

    const result: IProfileResult = {
        trace: calcService.getExecutionTrace(),
        peakHeapMB: 0,
        finalHeapMB: 0,
        wallMs: 0,
    };

    if (result.trace.length > 0) {
        result.peakHeapMB = Math.max(...result.trace.map((s) => s.heapMB));
        result.finalHeapMB = result.trace[result.trace.length - 1].heapMB;
        result.wallMs = result.trace[result.trace.length - 1].ms;
    }

    if (cpu) {
        const { profile } = await session!.post('Profiler.stop') as { profile: unknown };
        ensureProfileDir();
        const path = join(PROFILE_DIR, profileFilename(label, 'cpuprofile'));
        writeFileSync(path, JSON.stringify(profile));
        result.cpuProfilePath = path;
        await session!.post('Profiler.disable');
    }

    if (heap) {
        const { profile } = await session!.post('HeapProfiler.stopSampling') as { profile: unknown };
        ensureProfileDir();
        const path = join(PROFILE_DIR, profileFilename(label, 'heapprofile'));
        writeFileSync(path, JSON.stringify(profile));
        result.heapProfilePath = path;
        await session!.post('HeapProfiler.disable');
    }

    if (session) {
        session.disconnect();
    }

    if (result.trace.length > 0) {
        printTrace(result.trace);
    }

    if (result.cpuProfilePath) {
        console.warn(`CPU profile: ${result.cpuProfilePath}`);
    }
    if (result.heapProfilePath) {
        console.warn(`Heap profile: ${result.heapProfilePath}`);
    }

    return result;
}
