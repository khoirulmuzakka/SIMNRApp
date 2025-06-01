# SIMNRA COM Automation Interface in C++

## Description

This project provides a modern and lightweight C++ interface for automating the **SIMNRA** software using Microsoft's **COM (Component Object Model)** technology. SIMNRA is a widely used simulation tool for analyzing ion beam interactions, particularly for Rutherford Backscattering Spectrometry (RBS) and related techniques. 

The interface allows users to programmatically control SIMNRA, modify simulation parameters, and retrieve results without manual interaction.

The code includes utility functions for safely handling COM `VARIANT` types, invoking methods, and setting/retrieving properties from SIMNRA's COM-dispatched interfaces. A `SIMNRA` class wraps key modules like `App`, `Setup`, `Target`, `Calc`, and others, exposing convenient, high-level methods for common tasks such as:

- Loading and simulating a setup file.
- Reading and writing beam parameters (e.g., energy and energy spread).
- Managing target composition (layers, elements, thickness, concentrations).
- Triggering spectrum calculations.

This automation layer significantly enhances workflow efficiency for users performing batch simulations, parameter scans, or coupling SIMNRA with optimization frameworks and AI-based analysis pipelines.

## Key Features

- **COM Automation Wrapper**: Interfaces with all major SIMNRA modules (`App`, `Target`, `Setup`, `Calc`, etc.).
- **Robust COM Handling**: Includes exception-safe functions to invoke methods and retrieve property values using `IDispatch*`.
- **Type Conversion Utilities**: Template-based `variantTo<T>` and `InvokeMethod<T>` functions simplify `VARIANT` type handling.
- **Modular & Extendable**: Clean C++ design with support for extension and integration into larger software systems.

## Typical Use Cases

- Automating RBS simulations for multiple samples.
- Performing parameter sweeps or Monte Carlo simulations.
- Building custom graphical interfaces or optimization loops that drive SIMNRA.
- Integrating SIMNRA as a backend engine in larger data analysis workflows.

## Dependencies

- Windows OS (due to COM and `<windows.h>`)
- Microsoft COM infrastructure
- SIMNRA installed with a registered COM server (`SIMNRA.app`, `SIMNRA.setup`, etc.)

## License

This project is intended for academic and research use. Please ensure compliance with SIMNRA’s license terms for automation and interfacing.

