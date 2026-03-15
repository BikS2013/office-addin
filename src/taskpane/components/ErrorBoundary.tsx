import React from "react";
import {
  MessageBar,
  MessageBarBody,
  MessageBarActions,
  Button,
  makeStyles,
  tokens,
} from "@fluentui/react-components";

interface ErrorBoundaryProps {
  readonly children: React.ReactNode;
}

interface ErrorBoundaryState {
  readonly hasError: boolean;
  readonly error: Error | null;
}

const useStyles = makeStyles({
  container: {
    padding: tokens.spacingVerticalM,
  },
});

function ErrorFallback({
  error,
  onReset,
}: {
  readonly error: Error | null;
  readonly onReset: () => void;
}): React.ReactElement {
  const styles = useStyles();

  return (
    <div className={styles.container}>
      <MessageBar intent="error">
        <MessageBarBody>
          {error?.message ?? "An unexpected error occurred."}
        </MessageBarBody>
        <MessageBarActions>
          <Button size="small" onClick={onReset}>
            Try Again
          </Button>
        </MessageBarActions>
      </MessageBar>
    </div>
  );
}

export class ErrorBoundary extends React.Component<
  ErrorBoundaryProps,
  ErrorBoundaryState
> {
  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, errorInfo: React.ErrorInfo): void {
    console.error("ErrorBoundary caught an error:", error, errorInfo);
  }

  handleReset = (): void => {
    this.setState({ hasError: false, error: null });
  };

  render(): React.ReactNode {
    if (this.state.hasError) {
      return (
        <ErrorFallback error={this.state.error} onReset={this.handleReset} />
      );
    }

    return this.props.children;
  }
}
