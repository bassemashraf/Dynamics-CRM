/* eslint-disable */
/**
 * InitCache.ts
 * Singleton cache for bookable resource and booking status data.
 * Loaded once at component init — never during create flow.
 */

export class InitCache {
    private static get xrm(): Xrm.XrmStatic {
        return (window as any).Xrm;
    }

    private static _bookableResourceId: string | null = null;
    private static _bookingStatusId: string | null = null;
    private static _loaded = false;
    private static _loadPromise: Promise<void> | null = null;

    /**
     * Load bookable resource ID and scheduled booking status ID in parallel.
     * Safe to call multiple times — only fetches once.
     */
    static async load(userId: string): Promise<void> {
        if (this._loaded) return;

        // Deduplicate concurrent calls
        if (this._loadPromise) {
            return this._loadPromise;
        }

        this._loadPromise = this._doLoad(userId);
        await this._loadPromise;
    }

    private static async _doLoad(userId: string): Promise<void> {
        try {
            const [resourceResult, statusResult] = await Promise.all([
                this.xrm.WebApi.retrieveMultipleRecords(
                    'bookableresource',
                    `?$select=bookableresourceid&$filter=_userid_value eq ${userId} and resourcetype eq 3&$top=1`
                ),
                this.xrm.WebApi.retrieveMultipleRecords(
                    'bookingstatus',
                    `?$select=bookingstatusid&$filter=name eq 'Scheduled'&$top=1`
                )
            ]);

            this._bookableResourceId = resourceResult?.entities?.length > 0
                ? resourceResult.entities[0].bookableresourceid
                : null;

            this._bookingStatusId = statusResult?.entities?.length > 0
                ? statusResult.entities[0].bookingstatusid
                : null;

            alert(`InitCache: Bookable Resource: ${this._bookableResourceId ? 'Found' : 'Not Found'}, Booking Status: ${this._bookingStatusId ? 'Found' : 'Not Found'}`);

            this._loaded = true;
        } catch (error: any) {
            console.error('InitCache: Error loading cached data:', error);
            alert('InitCache: Error loading cached data: ' + (error?.message || error) + (error?.innerError ? '\nInner: ' + JSON.stringify(error.innerError) : ''));
            this._loaded = true; // Mark as loaded to prevent infinite retries
        } finally {
            this._loadPromise = null;
        }
    }

    /** Cached bookable resource ID for the current user */
    static get bookableResourceId(): string | null {
        return this._bookableResourceId;
    }

    /** Cached 'Scheduled' booking status ID */
    static get bookingStatusId(): string | null {
        return this._bookingStatusId;
    }

    /** True if user has a bookable resource record — replaces isInspector logic */
    static get hasBookableResource(): boolean {
        return this._bookableResourceId !== null;
    }

    /** Reset cache (e.g. on user change) */
    static reset(): void {
        this._bookableResourceId = null;
        this._bookingStatusId = null;
        this._loaded = false;
        this._loadPromise = null;
    }
}
