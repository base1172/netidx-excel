use fxhash::FxHashMap;
use netidx::{
    config::Config,
    path::Path,
    subscriber::{Dval, Subscriber, Value},
};
use tokio::sync::mpsc;

pub struct Setter {
    tx: mpsc::UnboundedSender<(Path, Value)>,
}

impl Setter {
    pub fn new() -> anyhow::Result<Self> {
        let cfg = Config::load_default()?;
        let (tx, mut rx) = mpsc::unbounded_channel::<(Path, Value)>();
        let rt = tokio::runtime::Builder::new_current_thread().enable_all().build()?;
        let desired_auth = cfg.default_auth();
        let subscriber =
            rt.block_on(async move { Subscriber::new(cfg, desired_auth) })?;
        std::thread::Builder::new().name("netidx-setter".into()).spawn(move || {
            // TODO: Add support for closing unused [Dval] by integrating with the RTD server. We need to marshal [Dval] ids to/from strings when a new path is subscribed or when an old path is dropped.
            // Paths that are already subscribed and still in use won't need to be marshaled.
            let mut subs: FxHashMap<Path, Dval> = FxHashMap::default();
            rt.block_on(async move {
                while let Some((path, value)) = rx.recv().await {
                    use std::collections::hash_map::Entry::*;
                    match subs.entry(path) {
                        Occupied(entry) => entry.get().write(value),
                        Vacant(vacant) => {
                            let path = vacant.key().clone();
                            vacant.insert(subscriber.subscribe(path)).write(value)
                        }
                    };
                }
            });
            log::error!("netidx-setter thread exited");
        })?;
        Ok(Setter { tx })
    }

    pub fn set(
        &self,
        path: Path,
        value: Value,
    ) -> Result<(), mpsc::error::SendError<(Path, Value)>> {
        self.tx.send((path, value))
    }
}
